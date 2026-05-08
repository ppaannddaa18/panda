"""主窗口"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, List
import threading

try:
    from tkinterdnd2 import TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    import tkinter as tk_base
    TkinterDnD = type("TkinterDnD", (), {"Tk": tk_base.Tk})

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from bank_reconciliation.models import Transaction, MatchResult, MatchStatus
from bank_reconciliation.core import BankStatementParser, AccountLedgerParser, MatchEngine, ResultExporter
from bank_reconciliation.config import AppConfig, BANK_TEMPLATES, ACCOUNT_TEMPLATES
from .file_loader import FileLoader


class BankReconciliationApp:
    """银行对账助手主应用"""

    def __init__(self, root: Optional[tk.Tk] = None):
        """
        初始化应用

        Args:
            root: Tkinter 根窗口
        """
        if root is None:
            root = TkinterDnD.Tk() if HAS_DND else tk.Tk()

        self.root = root
        self.config = AppConfig()

        # 数据
        self.bank_txns: List[Transaction] = []
        self.account_txns: List[Transaction] = []
        self.match_results: List[MatchResult] = []

        # 解析器和引擎
        self.bank_parser: Optional[BankStatementParser] = None
        self.account_parser: Optional[AccountLedgerParser] = None
        self.match_engine = MatchEngine(self.config)
        self.exporter = ResultExporter(self.config.output_dir)

        self._setup_window()
        self._setup_styles()
        self._setup_ui()

    def _setup_window(self):
        """设置窗口"""
        self.root.title("智能银行对账助手")
        self.root.geometry(f"{self.config.window_width}x{self.config.window_height}")
        self.root.minsize(800, 600)

    def _setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Arial", 14, "bold"))
        style.configure("Heading.TLabel", font=("Arial", 10, "bold"))
        style.configure("Success.TLabel", foreground="green")
        style.configure("Warning.TLabel", foreground="orange")
        style.configure("Action.TButton", font=("Arial", 10, "bold"))

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        self.main_container = ttk.Frame(self.root, padding=10)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(self.main_container)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="智能银行对账助手", style="Title.TLabel").pack(side=tk.LEFT)

        # 文件选择区域
        self._create_file_section()

        # 操作按钮区域
        self._create_action_section()

        # 结果预览区域
        self._create_result_section()

        # 状态栏
        self._create_status_section()

    def _create_file_section(self):
        """创建文件选择区域"""
        file_frame = ttk.Frame(self.main_container)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 银行流水加载器
        self.bank_loader = FileLoader(
            file_frame,
            "银行流水",
            on_file_loaded=self._on_bank_file_loaded
        )
        self.bank_loader.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # 账务数据加载器
        self.account_loader = FileLoader(
            file_frame,
            "账务数据",
            on_file_loaded=self._on_account_file_loaded
        )
        self.account_loader.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # 模板选择
        template_frame = ttk.LabelFrame(self.main_container, text=" 模板设置 ", padding=5)
        template_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(template_frame, text="银行模板:").pack(side=tk.LEFT, padx=(0, 5))
        self.bank_template_var = tk.StringVar(value="自定义")
        self.bank_template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.bank_template_var,
            values=list(BANK_TEMPLATES.keys()) + ["自定义"],
            state="readonly",
            width=12
        )
        self.bank_template_combo.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(template_frame, text="账务模板:").pack(side=tk.LEFT, padx=(0, 5))
        self.account_template_var = tk.StringVar(value="用友")
        self.account_template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.account_template_var,
            values=list(ACCOUNT_TEMPLATES.keys()) + ["自定义"],
            state="readonly",
            width=12
        )
        self.account_template_combo.pack(side=tk.LEFT)

    def _create_action_section(self):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(self.main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="开始匹配",
            style="Action.TButton",
            command=self._start_match
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            action_frame,
            text="导出结果",
            command=self._export_results
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            action_frame,
            text="清空数据",
            command=self._clear_data
        ).pack(side=tk.LEFT, padx=5)

    def _create_result_section(self):
        """创建结果预览区域"""
        result_frame = ttk.LabelFrame(self.main_container, text=" 匹配结果预览 ", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建 Treeview
        columns = ("status", "bank_date", "bank_amount", "account_date", "account_amount", "summary")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        self.result_tree.heading("status", text="状态")
        self.result_tree.heading("bank_date", text="银行日期")
        self.result_tree.heading("bank_amount", text="银行金额")
        self.result_tree.heading("account_date", text="账务日期")
        self.result_tree.heading("account_amount", text="账务金额")
        self.result_tree.heading("summary", text="摘要")

        self.result_tree.column("status", width=80)
        self.result_tree.column("bank_date", width=100)
        self.result_tree.column("bank_amount", width=100)
        self.result_tree.column("account_date", width=100)
        self.result_tree.column("account_amount", width=100)
        self.result_tree.column("summary", width=200)

        # 滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_status_section(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.main_container)
        status_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(
            status_frame,
            text="就绪 | 银行: 0 笔 | 账务: 0 笔 | 匹配: 0 笔",
            foreground="gray"
        )
        self.status_label.pack(side=tk.LEFT)

    def _on_bank_file_loaded(self, file_path: str):
        """银行文件加载回调"""
        try:
            template = self.bank_template_var.get()
            self.bank_parser = BankStatementParser(template_name=template if template != "自定义" else None)
            self.bank_txns = self.bank_parser.parse(file_path)
            self._update_status()
            messagebox.showinfo("成功", f"已加载 {len(self.bank_txns)} 笔银行流水记录")
        except Exception as e:
            messagebox.showerror("错误", f"加载银行流水失败: {e}")

    def _on_account_file_loaded(self, file_path: str):
        """账务文件加载回调"""
        try:
            template = self.account_template_var.get()
            self.account_parser = AccountLedgerParser(template_name=template if template != "自定义" else None)
            self.account_txns = self.account_parser.parse(file_path)
            self._update_status()
            messagebox.showinfo("成功", f"已加载 {len(self.account_txns)} 笔账务记录")
        except Exception as e:
            messagebox.showerror("错误", f"加载账务数据失败: {e}")

    def _start_match(self):
        """开始匹配"""
        if not self.bank_txns and not self.account_txns:
            messagebox.showwarning("警告", "请先加载银行流水和账务数据")
            return

        # 在后台线程执行匹配
        def do_match():
            self.match_results = self.match_engine.match(self.bank_txns, self.account_txns)
            self.root.after(0, self._on_match_complete)

        threading.Thread(target=do_match, daemon=True).start()
        self.status_label.config(text="正在匹配...")

    def _on_match_complete(self):
        """匹配完成回调"""
        self._display_results()
        self._update_status()

        # 统计
        matched = sum(1 for r in self.match_results if r.is_matched)
        total = len(self.match_results)
        rate = matched / total * 100 if total > 0 else 0

        messagebox.showinfo("匹配完成", f"匹配完成！\n匹配率: {rate:.1f}%\n已匹配: {matched} 笔")

    def _display_results(self):
        """显示匹配结果"""
        # 清空现有数据
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 添加结果
        for result in self.match_results:
            status = result.status_display
            bank_date = result.bank_txn.date.strftime("%Y-%m-%d") if result.bank_txn else "-"
            bank_amount = f"{float(result.bank_txn.amount):,.2f}" if result.bank_txn else "-"
            account_date = result.account_txn.date.strftime("%Y-%m-%d") if result.account_txn else "-"
            account_amount = f"{float(result.account_txn.amount):,.2f}" if result.account_txn else "-"
            summary = (result.bank_txn.summary[:20] if result.bank_txn else
                      result.account_txn.summary[:20] if result.account_txn else "-")

            tags = ("matched",) if result.is_matched else ("unmatched",)
            self.result_tree.insert("", tk.END, values=(
                status, bank_date, bank_amount, account_date, account_amount, summary
            ), tags=tags)

        # 设置标签样式
        self.result_tree.tag_configure("matched", background="#C6EFCE")
        self.result_tree.tag_configure("unmatched", background="#FFEB9C")

    def _export_results(self):
        """导出结果"""
        if not self.match_results:
            messagebox.showwarning("警告", "没有匹配结果可导出")
            return

        try:
            files = self.exporter.export_all(self.match_results)
            messagebox.showinfo(
                "导出成功",
                f"已导出以下文件:\n" + "\n".join(files.values())
            )
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def _clear_data(self):
        """清空数据"""
        self.bank_txns = []
        self.account_txns = []
        self.match_results = []
        self.bank_loader.clear()
        self.account_loader.clear()

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self._update_status()

    def _update_status(self):
        """更新状态栏"""
        matched = sum(1 for r in self.match_results if r.is_matched)
        self.status_label.config(
            text=f"就绪 | 银行: {len(self.bank_txns)} 笔 | 账务: {len(self.account_txns)} 笔 | 匹配: {matched} 笔"
        )

    def run(self):
        """运行应用"""
        self.root.mainloop()
