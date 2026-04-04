"""PDF 工具箱 - 合并 & 拆分

提供 PDF 文件的合并和拆分功能，支持拖拽操作。
"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
import ttkbootstrap as ttk
from ttkbootstrap.constants import *


# 配置常量
WINDOW_TITLE = "PDF 工具箱 - 合并 & 拆分"
WINDOW_SIZE = "700x580"
SUPPORTED_FORMAT = ".pdf"


class PDFToolApp:
    """PDF 工具箱应用主类"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(True, True)

        # 合并功能的数据
        self._merge_files: list[str] = []

        # 创建选项卡界面
        self._setup_ui()

    def _setup_ui(self) -> None:
        """初始化用户界面"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 合并 PDF 选项卡
        self.merge_frame = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(self.merge_frame, text="合并 PDF")
        self._setup_merge_tab()

        # 拆分 PDF 选项卡
        self.split_frame = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(self.split_frame, text="拆分 PDF")
        self._setup_split_tab()

    # =========================================================================
    # 合并功能
    # =========================================================================

    def _setup_merge_tab(self) -> None:
        """设置合并 PDF 选项卡"""
        frame = self.merge_frame

        # 标题
        ttk.Label(
            frame, text="合并多个 PDF 文件",
            font=("微软雅黑", 16, "bold"), bootstyle="primary"
        ).pack(pady=5)
        ttk.Label(
            frame, text="拖拽 PDF 文件到这里，或点击按钮添加",
            font=("微软雅黑", 9), bootstyle="secondary"
        ).pack(pady=5)

        # 文件列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.merge_listbox = tk.Listbox(
            list_frame, selectmode=tk.EXTENDED, height=10,
            font=("微软雅黑", 10)
        )
        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.merge_listbox.yview
        )
        self.merge_listbox.config(yscrollcommand=scrollbar.set)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 拖拽支持
        self.merge_listbox.drop_target_register(DND_FILES)
        self.merge_listbox.dnd_bind("<<Drop>>", self._on_merge_drop)
        self.merge_listbox.dnd_bind("<<DropEnter>>", self._on_merge_drag_enter)
        self.merge_listbox.dnd_bind("<<DropLeave>>", self._on_merge_drag_leave)

        # 操作按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(
            btn_frame, text="添加文件", bootstyle="info",
            command=self._merge_add_files
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame, text="删除选中", bootstyle="danger",
            command=self._merge_remove_selected
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame, text="清空列表", bootstyle="secondary",
            command=self._merge_clear_list
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_frame, text="合并 PDF", bootstyle="success",
            command=self._merge_pdfs
        ).pack(side=tk.RIGHT, padx=5)

        # 状态栏
        self.merge_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.merge_status, bootstyle="secondary").pack(
            side=tk.BOTTOM, fill=tk.X
        )

    def _on_merge_drag_enter(self, event) -> None:
        """拖拽进入时高亮显示"""
        self.merge_listbox.configure(bg="#e0f7fa")

    def _on_merge_drag_leave(self, event) -> None:
        """拖拽离开时恢复颜色"""
        self.merge_listbox.configure(bg="white")

    def _on_merge_drop(self, event) -> None:
        """处理拖拽放下事件"""
        self.merge_listbox.configure(bg="white")
        files = self.root.tk.splitlist(event.data)
        valid_files = [
            f for f in files
            if os.path.isfile(f) and f.lower().endswith(SUPPORTED_FORMAT)
        ]

        if valid_files:
            added_count = 0
            for filepath in valid_files:
                if filepath not in self._merge_files:
                    self.merge_listbox.insert(tk.END, os.path.basename(filepath))
                    self._merge_files.append(filepath)
                    added_count += 1
            self.merge_status.set(f"已添加 {added_count} 个文件")
        else:
            self.merge_status.set("无效的文件（仅支持 PDF）")

    def _merge_add_files(self) -> None:
        """通过文件对话框添加文件"""
        files = filedialog.askopenfilenames(
            title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")]
        )
        if files:
            added_count = 0
            for filepath in files:
                if filepath not in self._merge_files:
                    self.merge_listbox.insert(tk.END, os.path.basename(filepath))
                    self._merge_files.append(filepath)
                    added_count += 1
            self.merge_status.set(f"添加了 {added_count} 个文件")

    def _merge_remove_selected(self) -> None:
        """删除选中的文件"""
        selected = self.merge_listbox.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先选中要删除的文件")
            return

        for index in reversed(selected):
            self.merge_listbox.delete(index)
            del self._merge_files[index]
        self.merge_status.set("已删除选中文件")

    def _merge_clear_list(self) -> None:
        """清空文件列表"""
        if self._merge_files and messagebox.askyesno("确认", "确定清空文件列表吗？"):
            self.merge_listbox.delete(0, tk.END)
            self._merge_files.clear()
            self.merge_status.set("列表已清空")

    def _merge_pdfs(self) -> None:
        """执行 PDF 合并"""
        if not self._merge_files:
            messagebox.showwarning("警告", "请先添加文件！")
            return

        output_path = filedialog.asksaveasfilename(
            title="保存合并文件",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if not output_path:
            return

        try:
            self.merge_status.set("正在合并...")
            self.root.update_idletasks()

            merger = PdfMerger()
            for pdf_path in self._merge_files:
                merger.append(pdf_path)
            merger.write(output_path)
            merger.close()

            self.merge_status.set("合并成功！")
            messagebox.showinfo("成功", f"已保存至：\n{output_path}")
        except Exception as e:
            self.merge_status.set("合并失败")
            messagebox.showerror("错误", f"合并失败：\n{str(e)}")

    # =========================================================================
    # 拆分功能
    # =========================================================================

    def _setup_split_tab(self) -> None:
        """设置拆分 PDF 选项卡"""
        frame = self.split_frame

        # 标题
        ttk.Label(
            frame, text="拆分 PDF 文件",
            font=("微软雅黑", 16, "bold"), bootstyle="primary"
        ).pack(pady=5)
        ttk.Label(
            frame, text="支持两种拆分方式：指定页码 或 按每份页数",
            font=("微软雅黑", 9), bootstyle="secondary"
        ).pack(pady=5)

        # 拖拽区域
        self.drop_frame = ttk.Frame(frame, padding=10, relief="groove", borderwidth=2)
        self.drop_frame.pack(fill=tk.X, pady=10)

        self.drop_label = ttk.Label(
            self.drop_frame,
            text="拖拽 PDF 文件到这里\n或点击下方按钮选择",
            font=("微软雅黑", 10), bootstyle="secondary", anchor="center"
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_split_drop)
        self.drop_frame.dnd_bind("<<DropEnter>>", self._on_split_drag_enter)
        self.drop_frame.dnd_bind("<<DropLeave>>", self._on_split_drag_leave)

        # 选择文件按钮
        ttk.Button(
            frame, text="选择 PDF 文件", bootstyle="info",
            command=self._split_browse_file
        ).pack(pady=5)

        # 文件路径显示
        self.split_filepath = tk.StringVar()
        ttk.Entry(
            frame, textvariable=self.split_filepath,
            state="readonly", font=("微软雅黑", 9)
        ).pack(fill=tk.X, pady=5)

        # 拆分模式选择
        mode_frame = ttk.LabelFrame(frame, text="拆分模式")
        mode_frame.pack(fill=tk.X, pady=10, padx=10)

        self.split_mode = tk.StringVar(value="range")

        # 模式1：指定页码范围
        ttk.Radiobutton(
            mode_frame, text="1. 指定页码范围",
            variable=self.split_mode, value="range", bootstyle="primary"
        ).grid(row=0, column=0, sticky=tk.W, pady=2)

        self.range_var = tk.StringVar(value="1-3,5,7-10")
        ttk.Entry(
            mode_frame, textvariable=self.range_var,
            font=("微软雅黑", 9)
        ).grid(row=1, column=0, padx=20, pady=2, sticky=tk.EW)
        ttk.Label(
            mode_frame, text="格式：1,3-5,7",
            font=("微软雅黑", 8), bootstyle="secondary"
        ).grid(row=1, column=1, padx=5)

        # 模式2：按每份页数拆分
        ttk.Radiobutton(
            mode_frame, text="2. 按每份页数拆分",
            variable=self.split_mode, value="chunks", bootstyle="primary"
        ).grid(row=2, column=0, sticky=tk.W, pady=2)

        self.chunk_var = tk.IntVar(value=3)
        ttk.Spinbox(
            mode_frame, from_=1, to=100,
            textvariable=self.chunk_var, width=5
        ).grid(row=3, column=0, padx=20, pady=2, sticky=tk.W)
        ttk.Label(
            mode_frame, text="页/份",
            font=("微软雅黑", 9)
        ).grid(row=3, column=0, padx=70, pady=2, sticky=tk.W)

        mode_frame.columnconfigure(0, weight=1)

        # 开始拆分按钮
        ttk.Button(
            frame, text="开始拆分", bootstyle="warning",
            command=self._split_pdf
        ).pack(pady=20)

        # 状态栏
        self.split_status = tk.StringVar(value="就绪")
        ttk.Label(frame, textvariable=self.split_status, bootstyle="secondary").pack(
            side=tk.BOTTOM, fill=tk.X
        )

    def _on_split_drag_enter(self, event) -> None:
        """拖拽进入时高亮显示"""
        self.drop_frame.configure(bootstyle="success")
        self.drop_label.configure(bootstyle="success")

    def _on_split_drag_leave(self, event) -> None:
        """拖拽离开时恢复颜色"""
        self.drop_frame.configure(bootstyle="default")
        self.drop_label.configure(bootstyle="secondary")

    def _on_split_drop(self, event) -> None:
        """处理拖拽放下事件"""
        self._on_split_drag_leave(event)
        files = self.root.tk.splitlist(event.data)
        pdf_files = [
            f for f in files
            if os.path.isfile(f) and f.lower().endswith(SUPPORTED_FORMAT)
        ]

        if pdf_files:
            self.split_filepath.set(pdf_files[0])
            self.split_status.set(f"已选择：{os.path.basename(pdf_files[0])}")
        else:
            self.split_status.set("无效的文件（仅支持 PDF）")

    def _split_browse_file(self) -> None:
        """通过文件对话框选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择要拆分的 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if file_path:
            self.split_filepath.set(file_path)
            self.split_status.set(f"已选择：{os.path.basename(file_path)}")

    def _parse_page_ranges(self, page_str: str, total_pages: int) -> Optional[list[int]]:
        """解析页码字符串，返回页码列表（从1开始）

        Args:
            page_str: 页码字符串，如 "1-3,5,7-10"
            total_pages: PDF 总页数

        Returns:
            排序后的页码列表，解析失败返回 None
        """
        page_str = re.sub(r"\s+", "", page_str)
        if not page_str:
            return None

        result = set()
        parts = page_str.split(",")

        for part in parts:
            part = part.strip()
            if not part:
                continue

            if "-" in part:
                try:
                    start_str, end_str = part.split("-", 1)
                    start, end = int(start_str), int(end_str)
                    if start > end:
                        start, end = end, start
                    start = max(1, start)
                    end = min(total_pages, end)
                    result.update(range(start, end + 1))
                except (ValueError, IndexError):
                    return None
            else:
                try:
                    page = int(part)
                    if 1 <= page <= total_pages:
                        result.add(page)
                except ValueError:
                    return None

        return sorted(result) if result else None

    def _split_pdf(self) -> None:
        """执行 PDF 拆分"""
        file_path = self.split_filepath.get().strip()
        if not file_path or not os.path.isfile(file_path):
            messagebox.showwarning("警告", "请先选择有效的 PDF 文件！")
            return

        try:
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            self.split_status.set(f"文档共 {total_pages} 页，正在处理...")
            self.root.update_idletasks()

            mode = self.split_mode.get()
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_dir = filedialog.askdirectory(title="选择保存文件夹")
            if not output_dir:
                return

            if mode == "range":
                self._split_by_range(reader, total_pages, base_name, output_dir)
            elif mode == "chunks":
                self._split_by_chunks(reader, total_pages, base_name, output_dir)

        except Exception as e:
            self.split_status.set("拆分失败")
            messagebox.showerror("错误", f"拆分失败：\n{str(e)}")

    def _split_by_range(
        self, reader: PdfReader, total_pages: int,
        base_name: str, output_dir: str
    ) -> None:
        """按指定页码范围拆分"""
        page_input = self.range_var.get().strip()
        if not page_input:
            messagebox.showwarning("警告", "请输入页码范围！")
            return

        page_numbers = self._parse_page_ranges(page_input, total_pages)
        if not page_numbers:
            messagebox.showerror("错误", "页码格式不正确！\n示例：1,3-5,7")
            return

        output_file = os.path.join(output_dir, f"{base_name}_extract.pdf")
        writer = PdfWriter()
        for page_num in page_numbers:
            writer.add_page(reader.pages[page_num - 1])

        with open(output_file, "wb") as f:
            writer.write(f)

        self.split_status.set("拆分完成（指定页码）")
        messagebox.showinfo("成功", f"已保存为：\n{output_file}")

    def _split_by_chunks(
        self, reader: PdfReader, total_pages: int,
        base_name: str, output_dir: str
    ) -> None:
        """按每份页数拆分"""
        chunk_size = self.chunk_var.get()
        if chunk_size < 1:
            messagebox.showerror("错误", "每份页数必须 >= 1")
            return

        part = 1
        for start in range(0, total_pages, chunk_size):
            end = min(start + chunk_size, total_pages)
            writer = PdfWriter()
            for i in range(start, end):
                writer.add_page(reader.pages[i])

            output_file = os.path.join(output_dir, f"{base_name}_part_{part}.pdf")
            with open(output_file, "wb") as f:
                writer.write(f)
            part += 1

        self.split_status.set(f"拆分完成（{part - 1} 个文件）")
        messagebox.showinfo("成功", f"已拆分为 {part - 1} 个文件，保存在：\n{output_dir}")


def main():
    """程序入口"""
    root = TkinterDnD.Tk()
    app = PDFToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()