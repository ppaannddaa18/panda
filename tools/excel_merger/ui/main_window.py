"""主窗口"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, List
from pathlib import Path
import threading

try:
    from tkinterdnd2 import TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    import tkinter as tk_base
    TkinterDnD = type("TkinterDnD", (), {"Tk": tk_base.Tk})

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from excel_merger.models import ParsedFile, MergeResult
from excel_merger.core import ExcelParser, AccountMatcher, ResultExporter
from excel_merger.config import AppConfig
from .mapping_editor import MappingEditor


class ExcelMergerApp:
    """Excel 多表合并工具主应用"""

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
        self.folder_path: str = ""
        self.parsed_files: List[ParsedFile] = []
        self.merge_result: Optional[MergeResult] = None

        # 核心组件
        self.parser = ExcelParser(self.config)
        self.matcher = AccountMatcher()
        self.exporter = ResultExporter(self.config.output_dir)

        self._setup_window()
        self._setup_styles()
        self._setup_ui()

    def _setup_window(self):
        """设置窗口"""
        self.root.title("Excel 多表合并工具箱")
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
        style.configure("Action.TButton", font=("Arial", 10, "bold"))

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        self.main_container = ttk.Frame(self.root, padding=10)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(self.main_container)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="Excel 多表合并工具箱", style="Title.TLabel").pack(side=tk.LEFT)

        # 文件夹选择区域
        self._create_folder_section()

        # 文件预览和设置区域
        self._create_preview_section()

        # 结果预览区域
        self._create_result_section()

        # 操作按钮区域
        self._create_action_section()

        # 状态栏
        self._create_status_section()

    def _create_folder_section(self):
        """创建文件夹选择区域"""
        folder_frame = ttk.LabelFrame(self.main_container, text=" 源文件文件夹 ", padding=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件夹路径
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X)

        self.folder_entry = ttk.Entry(path_frame, width=60)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(
            path_frame,
            text="选择文件夹",
            command=self._select_folder
        ).pack(side=tk.LEFT)

        # 文件计数
        self.file_count_label = ttk.Label(folder_frame, text="未选择文件夹", foreground="gray")
        self.file_count_label.pack(anchor=tk.W, pady=(5, 0))

    def _create_preview_section(self):
        """创建预览区域"""
        preview_frame = ttk.Frame(self.main_container)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 左侧：文件列表
        left_frame = ttk.LabelFrame(preview_frame, text=" 文件列表 ", padding=5)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # 文件 Treeview
        columns = ("filename", "company", "records", "status")
        self.file_tree = ttk.Treeview(left_frame, columns=columns, show="headings", height=8)

        self.file_tree.heading("filename", text="文件名")
        self.file_tree.heading("company", text="公司")
        self.file_tree.heading("records", text="记录数")
        self.file_tree.heading("status", text="状态")

        self.file_tree.column("filename", width=200)
        self.file_tree.column("company", width=100)
        self.file_tree.column("records", width=80)
        self.file_tree.column("status", width=80)

        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧：设置
        right_frame = ttk.LabelFrame(preview_frame, text=" 合并设置 ", padding=5)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

        ttk.Button(
            right_frame,
            text="编辑科目映射...",
            command=self._open_mapping_editor
        ).pack(fill=tk.X, pady=5)

    def _create_result_section(self):
        """创建结果预览区域"""
        result_frame = ttk.LabelFrame(self.main_container, text=" 合并结果预览 ", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 结果 Treeview
        columns = ("account", "company", "month", "amount")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=10)

        self.result_tree.heading("account", text="科目")
        self.result_tree.heading("company", text="公司")
        self.result_tree.heading("month", text="月份")
        self.result_tree.heading("amount", text="金额")

        self.result_tree.column("account", width=200)
        self.result_tree.column("company", width=100)
        self.result_tree.column("month", width=80)
        self.result_tree.column("amount", width=120)

        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_action_section(self):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(self.main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="开始合并",
            style="Action.TButton",
            command=self._start_merge
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

    def _create_status_section(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.main_container)
        status_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(
            status_frame,
            text="就绪 | 请选择包含 Excel 文件的文件夹",
            foreground="gray"
        )
        self.status_label.pack(side=tk.LEFT)

    def _select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if folder:
            self.folder_path = folder
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            self._load_files()

    def _load_files(self):
        """加载文件"""
        if not self.folder_path:
            return

        self.status_label.config(text="正在加载文件...")
        self.root.update()

        # 解析文件
        self.parsed_files = self.parser.parse_folder(self.folder_path)

        # 更新文件列表
        self._update_file_tree()

        # 更新状态
        total_records = sum(len(f.all_records) for f in self.parsed_files)
        success_count = sum(1 for f in self.parsed_files if not f.parse_errors)

        self.file_count_label.config(
            text=f"已加载 {len(self.parsed_files)} 个文件，{total_records} 条记录（{success_count} 个成功）"
        )
        self.status_label.config(text=f"就绪 | 已加载 {len(self.parsed_files)} 个文件")

    def _update_file_tree(self):
        """更新文件列表"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        for parsed in self.parsed_files:
            filename = Path(parsed.file_path).name
            company = parsed.company
            records = len(parsed.all_records)
            status = "成功" if not parsed.parse_errors else "有错误"

            self.file_tree.insert("", tk.END, values=(filename, company, records, status))

    def _start_merge(self):
        """开始合并"""
        if not self.parsed_files:
            messagebox.showwarning("警告", "请先选择包含 Excel 文件的文件夹")
            return

        self.status_label.config(text="正在合并...")
        self.root.update()

        # 在后台线程执行
        def do_merge():
            self.merge_result = self.matcher.match(self.parsed_files)
            self.root.after(0, self._on_merge_complete)

        threading.Thread(target=do_merge, daemon=True).start()

    def _on_merge_complete(self):
        """合并完成回调"""
        self._display_results()
        self._update_status()

        unmatched_count = len(self.merge_result.unmatched_accounts)
        messagebox.showinfo(
            "合并完成",
            f"合并完成！\n"
            f"总记录数: {self.merge_result.record_count}\n"
            f"公司数量: {self.merge_result.company_count}\n"
            f"未匹配科目: {unmatched_count}"
        )

    def _display_results(self):
        """显示合并结果"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        if not self.merge_result:
            return

        # 只显示前 100 条记录
        for record in self.merge_result.records[:100]:
            self.result_tree.insert("", tk.END, values=(
                record.standardized_name,
                record.company,
                f"{record.month}月",
                f"{float(record.amount):,.2f}"
            ))

    def _export_results(self):
        """导出结果"""
        if not self.merge_result:
            messagebox.showwarning("警告", "请先执行合并操作")
            return

        try:
            files = self.exporter.export_all(self.merge_result)
            messagebox.showinfo(
                "导出成功",
                f"已导出以下文件:\n" + "\n".join(files.values())
            )
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def _clear_data(self):
        """清空数据"""
        self.folder_path = ""
        self.parsed_files = []
        self.merge_result = None

        self.folder_entry.delete(0, tk.END)
        self.file_count_label.config(text="未选择文件夹")

        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.status_label.config(text="就绪 | 请选择包含 Excel 文件的文件夹")

    def _update_status(self):
        """更新状态栏"""
        if self.merge_result:
            self.status_label.config(
                text=f"就绪 | 记录: {self.merge_result.record_count} | "
                     f"公司: {self.merge_result.company_count} | "
                     f"未匹配科目: {len(self.merge_result.unmatched_accounts)}"
            )

    def _open_mapping_editor(self):
        """打开科目映射编辑器"""
        MappingEditor(self.root, self.matcher)

    def run(self):
        """运行应用"""
        self.root.mainloop()
