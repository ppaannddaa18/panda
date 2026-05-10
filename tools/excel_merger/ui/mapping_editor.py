"""科目映射编辑器"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from excel_merger.core import AccountMatcher


class MappingEditor:
    """科目映射编辑器对话框"""

    def __init__(self, parent: tk.Tk, matcher: "AccountMatcher"):
        """
        初始化编辑器

        Args:
            parent: 父窗口
            matcher: 科目匹配器
        """
        self.matcher = matcher

        # 创建对话框
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("科目映射编辑器")
        self.dialog.geometry("600x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._setup_ui()
        self._load_mappings()

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 映射列表
        list_frame = ttk.LabelFrame(main_frame, text=" 当前映射 ", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview
        columns = ("original", "standardized")
        self.mapping_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)

        self.mapping_tree.heading("original", text="原始科目名称")
        self.mapping_tree.heading("standardized", text="标准化科目名称")

        self.mapping_tree.column("original", width=250)
        self.mapping_tree.column("standardized", width=250)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        self.mapping_tree.configure(yscrollcommand=scrollbar.set)

        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 添加/编辑区域
        edit_frame = ttk.LabelFrame(main_frame, text=" 添加/编辑映射 ", padding=5)
        edit_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(edit_frame, text="原始科目:").grid(row=0, column=0, padx=5, pady=5)
        self.original_entry = ttk.Entry(edit_frame, width=30)
        self.original_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(edit_frame, text="标准科目:").grid(row=0, column=2, padx=5, pady=5)
        self.standardized_entry = ttk.Entry(edit_frame, width=30)
        self.standardized_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Button(edit_frame, text="添加", command=self._add_mapping).grid(row=0, column=4, padx=5, pady=5)
        ttk.Button(edit_frame, text="删除", command=self._delete_mapping).grid(row=0, column=5, padx=5, pady=5)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="保存", command=self._save_mappings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)

        # 绑定选择事件
        self.mapping_tree.bind("<<TreeviewSelect>>", self._on_select)

    def _load_mappings(self):
        """加载映射列表"""
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)

        for original, standardized in self.matcher.account_mapping.items():
            self.mapping_tree.insert("", tk.END, values=(original, standardized))

    def _add_mapping(self):
        """添加映射"""
        original = self.original_entry.get().strip()
        standardized = self.standardized_entry.get().strip()

        if not original or not standardized:
            messagebox.showwarning("警告", "请输入原始科目名称和标准科目名称")
            return

        self.matcher.add_mapping(original, standardized)
        self._load_mappings()

        self.original_entry.delete(0, tk.END)
        self.standardized_entry.delete(0, tk.END)

    def _delete_mapping(self):
        """删除映射"""
        selected = self.mapping_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请选择要删除的映射")
            return

        item = selected[0]
        values = self.mapping_tree.item(item, "values")
        original = values[0]

        self.matcher.remove_mapping(original)
        self._load_mappings()

    def _on_select(self, event):
        """选择映射项"""
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, "values")

            self.original_entry.delete(0, tk.END)
            self.original_entry.insert(0, values[0])

            self.standardized_entry.delete(0, tk.END)
            self.standardized_entry.insert(0, values[1])

    def _save_mappings(self):
        """保存映射"""
        try:
            self.matcher.save_mapping("output/account_mapping.json")
            messagebox.showinfo("成功", "科目映射已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
