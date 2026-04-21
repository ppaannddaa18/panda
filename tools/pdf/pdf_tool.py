"""PDF 工具箱 - 合并 & 拆分 & 旋转

提供 PDF 文件的合并、拆分和旋转功能，支持拖拽操作。
"""

import os
import queue
import re
import threading
import tkinter as tk
from collections import OrderedDict
from datetime import datetime
from tkinter import filedialog, messagebox
from typing import Callable, Optional

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from tkinterdnd2 import DND_FILES, TkinterDnD
import ttkbootstrap as ttk

try:
    import fitz
    from PIL import Image, ImageTk
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


# 常量配置
WINDOW_TITLE = "PDF 工具箱 - 合并 & 拆分"
WINDOW_SIZE = "960x720"
WINDOW_MIN_SIZE = (700, 550)
SUPPORTED_FORMAT = ".pdf"
DRAG_BG_HIGHLIGHT = "#e0f7fa"
DRAG_BG_NORMAL = "white"

# 缓存配置
_MAX_CACHE_PAGES = 50
_THUMBNAIL_CACHE_SIZE = 100

# 缩略图配置
_THUMBNAIL_SIZE = (100, 140)
_THUMBNAIL_COLUMNS = 4


def _format_size(size_bytes: int) -> str:
    if size_bytes >= 1_048_576:
        return f"{size_bytes / 1_048_576:.1f} MB"
    return f"{size_bytes / 1024:.0f} KB"


def _status_msg(text: str) -> str:
    now = datetime.now().strftime("%H:%M:%S")
    return f"[{now}] {text}"


class _PageCache:
    """页面渲染缓存管理器，使用 LRU 策略"""

    def __init__(self, max_pages: int = _MAX_CACHE_PAGES):
        self._cache: OrderedDict = OrderedDict()
        self._max_pages = max_pages

    def get(self, filepath: str, page_num: int, dpi: int = 100):
        """获取缓存的页面图像"""
        key = (filepath, dpi)
        if key in self._cache and page_num in self._cache[key]:
            # 移动到末尾表示最近使用
            self._cache.move_to_end(key)
            return self._cache[key][page_num]
        return None

    def set(self, filepath: str, page_num: int, image, dpi: int = 100):
        """缓存页面图像"""
        key = (filepath, dpi)
        if key not in self._cache:
            self._cache[key] = OrderedDict()

        self._cache[key][page_num] = image
        self._cache.move_to_end(key)
        self._evict()

    def _evict(self):
        """LRU 淘汰策略"""
        total = sum(len(v) for v in self._cache.values())
        while total > self._max_pages and self._cache:
            oldest_key = next(iter(self._cache))
            oldest_pages = self._cache[oldest_key]
            if oldest_pages:
                oldest_pages.popitem(last=False)
                total -= 1
            if not oldest_pages:
                self._cache.pop(oldest_key)

    def clear(self):
        """清空缓存"""
        self._cache.clear()

    def remove_file(self, filepath: str):
        """移除指定文件的所有缓存"""
        keys_to_remove = [k for k in self._cache if k[0] == filepath]
        for k in keys_to_remove:
            self._cache.pop(k, None)


# 全局缓存实例
_page_cache = _PageCache()


class PDFToolApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)

        # 任务队列与线程安全通信
        self._task_queue = queue.Queue()

        # 合并数据
        self._merge_files: list[str] = []
        self._merge_page_counts: dict[str, int] = {}
        self._merge_busy = False

        # 拆分/旋转 busy 锁
        self._split_busy = False
        self._rotate_busy = False

        self._setup_ui()
        self._start_queue_poller()

    def __del__(self):
        """清理 PdfReader 缓存"""
        for reader in getattr(self, "_open_readers", []):
            try:
                reader.stream.close()
            except Exception:
                pass

    # =========================================================================
    # 队列轮询器
    # =========================================================================

    def _start_queue_poller(self) -> None:
        self._poll_queue()

    def _poll_queue(self) -> None:
        while not self._task_queue.empty():
            try:
                callback = self._task_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception as e:
                self.root.after(
                    0, lambda err=e: messagebox.showerror("错误", f"内部错误：{err}")
                )
        self.root.after(100, self._poll_queue)

    def _dispatch(self, fn, *args, **kwargs):
        if args or kwargs:
            self._task_queue.put(lambda: fn(*args, **kwargs))
        else:
            self._task_queue.put(fn)

    # =========================================================================
    # UI 初始化
    # =========================================================================

    def _setup_ui(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.merge_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.merge_frame, text="合并 PDF")
        self._setup_merge_tab()

        self.split_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.split_frame, text="拆分 PDF")
        self._setup_split_tab()

        self.rotate_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.rotate_frame, text="旋转页面")
        self._setup_rotate_tab()

        self.root.bind("<Control-a>", self._on_ctrl_a)
        self.root.bind("<Control-m>", lambda _: self._execute_current_tab())
        self.root.bind("<Control-s>", lambda _: self._execute_current_tab())

    # =========================================================================
    # 合并功能
    # =========================================================================

    def _setup_merge_tab(self) -> None:
        frame = self.merge_frame

        # 左右分栏容器
        left_frame = ttk.Frame(frame, width=380)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.pack_propagate(False)

        right_frame = ttk.Frame(frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # === 左侧：文件选择与设置 ===
        # 标题区域
        title_frame = ttk.Frame(left_frame)
        title_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            title_frame, text="合并多个 PDF 文件",
            font=("微软雅黑", 12, "bold"), bootstyle="primary"
        ).pack()
        ttk.Label(
            title_frame, text="拖拽或添加文件，按顺序合并",
            font=("微软雅黑", 9), bootstyle="secondary"
        ).pack()

        # 文件列表（带边框，拖拽时高亮）
        list_container = ttk.LabelFrame(left_frame, text="文件列表")  # type: ignore[attr-defined]
        list_container.pack(fill=tk.BOTH, expand=True, pady=5)

        list_border = ttk.Frame(list_container, bootstyle="default", relief="solid", borderwidth=1)
        list_border.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._list_border = list_border

        self.merge_listbox = tk.Listbox(
            list_border, selectmode=tk.EXTENDED, height=8,
            font=("微软雅黑", 9), bg="#fafafa", selectbackground="#4a90d9"
        )
        scrollbar = ttk.Scrollbar(
            list_border, orient="vertical", command=self.merge_listbox.yview
        )
        self.merge_listbox.config(yscrollcommand=scrollbar.set)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.merge_listbox.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        self.merge_listbox.dnd_bind("<<Drop>>", self._on_merge_drop)  # type: ignore[attr-defined]
        self.merge_listbox.dnd_bind("<<DropEnter>>", self._on_merge_drag_enter)  # type: ignore[attr-defined]
        self.merge_listbox.dnd_bind("<<DropLeave>>", self._on_merge_drag_leave)  # type: ignore[attr-defined]
        self.merge_listbox.bind("<Motion>", self._on_merge_hover)
        self.merge_listbox.bind("<Leave>", lambda _: self.merge_status.set(_status_msg(_current_merge_count(self))))
        self.merge_listbox.bind("<Delete>", lambda _: self._merge_remove_selected())
        self.merge_listbox.bind("<Return>", lambda _: self._merge_pdfs())
        self.merge_listbox.bind("<<ListboxSelect>>", lambda _: self._on_merge_selection_for_preview())

        # 按钮区域
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=8)

        btn_left = ttk.Frame(btn_frame)
        btn_left.pack(side=tk.LEFT)
        ttk.Button(
            btn_left, text="添加文件", bootstyle="info",
            command=self._merge_add_files, width=10
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_left, text="添加文件夹", bootstyle="info-outline",
            command=self._merge_add_folder, width=10
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_left, text="删除选中", bootstyle="danger",
            command=self._merge_remove_selected, width=10
        ).pack(side=tk.LEFT, padx=2)

        btn_right = ttk.Frame(btn_frame)
        btn_right.pack(side=tk.RIGHT)
        ttk.Button(
            btn_right, text="上移", bootstyle="secondary-outline",
            command=self._merge_move_up, width=6
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_right, text="下移", bootstyle="secondary-outline",
            command=self._merge_move_down, width=6
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_right, text="清空", bootstyle="secondary-outline",
            command=self._merge_clear_list, width=6
        ).pack(side=tk.LEFT, padx=2)

        # 合并预览信息
        preview_frame = ttk.LabelFrame(left_frame, text="合并预览")
        preview_frame.pack(fill=tk.X, pady=5)
        self.merge_preview_info = tk.StringVar(value="")
        ttk.Label(
            preview_frame, textvariable=self.merge_preview_info,
            bootstyle="info", font=("微软雅黑", 9)
        ).pack(pady=5)

        # 操作按钮
        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=8)
        ttk.Button(
            left_frame, text="合并 PDF", bootstyle="success",
            command=self._merge_pdfs, width=20
        ).pack(pady=5)

        # 进度条区域
        self.merge_progress_area = ttk.Frame(left_frame)
        self.merge_progress_bar = ttk.Progressbar(
            self.merge_progress_area, bootstyle="success-striped")
        self.merge_progress_bar.pack(fill=tk.X)
        self.merge_progress_label = ttk.Label(
            self.merge_progress_area, text="", bootstyle="secondary", font=("微软雅黑", 8))
        self.merge_progress_label.pack(fill=tk.X)
        self.merge_progress_area.pack_forget()

        # === 右侧：PDF 预览 ===
        self.merge_preview = _PDFPreviewPanel(right_frame, self, max_display_width=450, dpi=100)
        self.merge_preview.frame.pack(fill=tk.BOTH, expand=True)

        # 状态栏（底部）
        self.merge_status = tk.StringVar(value=_status_msg("就绪"))
        ttk.Label(frame, textvariable=self.merge_status, bootstyle="secondary").pack(
            side=tk.BOTTOM, fill=tk.X
        )

    def _on_merge_drag_enter(self, _event) -> None:
        self._list_border.configure(bootstyle="success")

    def _on_merge_drag_leave(self, _event) -> None:
        self._list_border.configure(bootstyle="default")

    def _on_merge_drop(self, event) -> None:
        self._list_border.configure(bootstyle="default")
        files = self.root.tk.splitlist(event.data)
        count = 0
        for f in files:
            f = os.path.abspath(f)
            if os.path.isfile(f) and f.lower().endswith(SUPPORTED_FORMAT):
                if f not in self._merge_files:
                    self._merge_files.append(f)
                    self.merge_listbox.insert(tk.END, _build_display(f))
                    self._cache_pages(f)
                    count += 1
            elif os.path.isdir(f):
                count += self._add_folder_to_merge(f)
        if count > 0:
            self.merge_status.set(_status_msg(f"已添加 {count} 个文件"))
            _update_merge_count(self)
            _update_merge_preview(self)
        else:
            self.merge_status.set(_status_msg("未找到有效的 PDF 文件"))

    def _on_merge_hover(self, event) -> None:
        idx = self.merge_listbox.nearest(event.y)
        if 0 <= idx < len(self._merge_files):
            self.merge_status.set(self._merge_files[idx])

    def _on_merge_selection_for_preview(self) -> None:
        sel = self.merge_listbox.curselection()
        if not sel:
            self.merge_preview.clear()
            return
        fp = self._merge_files[sel[0]]
        self.merge_preview.set_file(fp)

    def _cache_pages(self, filepath: str) -> int:
        if filepath in self._merge_page_counts:
            return self._merge_page_counts[filepath]
        try:
            r = PdfReader(filepath)
            n = len(r.pages)
            r.stream.close()
        except Exception:
            n = 0
        self._merge_page_counts[filepath] = n
        return n

    def _rebuild_listbox(self, sel_indices: Optional[set[int]] = None) -> None:
        self.merge_listbox.delete(0, tk.END)
        for fp in self._merge_files:
            self.merge_listbox.insert(tk.END, _build_display(fp))
        if sel_indices:
            for i in sorted(sel_indices):
                if 0 <= i < len(self._merge_files):
                    self.merge_listbox.selection_set(i)

    def _merge_add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")]
        )
        if files:
            count = 0
            for fp in files:
                fp = os.path.abspath(fp)
                if fp not in self._merge_files:
                    self._merge_files.append(fp)
                    self.merge_listbox.insert(tk.END, _build_display(fp))
                    self._cache_pages(fp)
                    count += 1
            self.merge_status.set(_status_msg(f"添加了 {count} 个文件"))
            _update_merge_count(self)
            _update_merge_preview(self)

    def _merge_add_folder(self) -> None:
        """添加整个文件夹中的 PDF 文件"""
        folder = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if not folder:
            return

        count = self._add_folder_to_merge(folder)
        if count > 0:
            self.merge_status.set(_status_msg(f"从文件夹添加了 {count} 个文件"))
            _update_merge_count(self)
            _update_merge_preview(self)
        else:
            messagebox.showinfo("提示", "该文件夹中没有 PDF 文件")

    def _add_folder_to_merge(self, folder: str) -> int:
        """递归查找文件夹中的 PDF 并添加到合并列表"""
        pdf_files = []
        for root_dir, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(SUPPORTED_FORMAT):
                    pdf_files.append(os.path.join(root_dir, f))

        pdf_files.sort(key=lambda x: os.path.basename(x).lower())

        count = 0
        for fp in pdf_files:
            fp = os.path.abspath(fp)
            if fp not in self._merge_files:
                self._merge_files.append(fp)
                self.merge_listbox.insert(tk.END, _build_display(fp))
                self._cache_pages(fp)
                count += 1

        return count

    def _merge_remove_selected(self) -> None:
        selected = self.merge_listbox.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先选中要删除的文件")
            return
        for i in reversed(selected):
            fp = self._merge_files[i]
            self.merge_listbox.delete(i)
            del self._merge_files[i]
            self._merge_page_counts.pop(fp, None)
        if self._merge_files:
            self.merge_listbox.selection_set(0)
            self._on_merge_selection_for_preview()
        else:
            self.merge_preview.clear()
        self.merge_status.set(_status_msg("已删除选中文件"))
        _update_merge_count(self)

    def _merge_clear_list(self) -> None:
        if not self._merge_files:
            return
        if messagebox.askyesno("确认", "确定清空文件列表吗？"):
            self.merge_listbox.delete(0, tk.END)
            self._merge_files.clear()
            self._merge_page_counts.clear()
            self.merge_preview.clear()
            self.merge_status.set(_status_msg("列表已清空"))

    def _merge_move_up(self) -> None:
        self._merge_move(-1)

    def _merge_move_down(self) -> None:
        self._merge_move(1)

    def _merge_move(self, direction: int) -> None:
        sel = list(self.merge_listbox.curselection())
        if not sel or not self._merge_files:
            return
        n = len(self._merge_files)
        if direction == -1 and 0 in sel:
            return
        if direction == 1 and (n - 1) in sel:
            return
        indices = sel if direction == -1 else list(reversed(sel))
        swap_set = set()
        for i in indices:
            j = i + direction
            if 0 <= j < n and i not in swap_set and j not in swap_set:
                self._merge_files[i], self._merge_files[j] = self._merge_files[j], self._merge_files[i]
                swap_set.add(i)
                swap_set.add(j)
        self._rebuild_listbox(set(sel))
        self._on_merge_selection_for_preview()
        _update_merge_count(self)


    def _merge_pdfs(self) -> None:
        if not self._merge_files:
            messagebox.showwarning("警告", "请先添加文件！")
            return
        if self._merge_busy:
            messagebox.showinfo("提示", "正在合并中，请稍候...")
            return

        output_path = filedialog.asksaveasfilename(
            title="保存合并文件",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if not output_path:
            return

        self._merge_busy = True
        file_list = list(self._merge_files)
        thread = threading.Thread(
            target=self._merge_worker, args=(output_path, file_list), daemon=True
        )
        thread.start()

    def _merge_worker(self, output_path: str, file_paths: list[str]) -> None:
        try:
            self._dispatch(self._set_merge_progress_visible, True)
            self._dispatch(lambda: self.merge_status.set(_status_msg("正在合并...")))

            merger = PdfMerger()
            total = len(file_paths)
            for idx, path in enumerate(file_paths, 1):
                merger.append(path)
                pct = int(idx / total * 60)
                self._dispatch(lambda p=pct: self._set_merge_progress(p))

            self._dispatch(lambda: self._set_merge_progress(75))
            merger.write(output_path)
            merger.close()
            self._dispatch(lambda: self._set_merge_progress(100))
            self._dispatch(lambda: self._merge_complete(output_path))
        except Exception as e:
            self._dispatch(lambda err=str(e): self._merge_fail(err))
        finally:
            self._dispatch(lambda: self._set_merge_busy(False))
            self._dispatch(self._set_merge_progress_visible, False)

    def _merge_complete(self, output_path: str) -> None:
        self.merge_status.set(_status_msg("合并成功！"))
        messagebox.showinfo("成功", f"已保存至：\n{output_path}")

    def _merge_fail(self, error: str) -> None:
        self.merge_status.set(_status_msg("合并失败"))
        messagebox.showerror("错误", f"合并失败：\n{error}")

    def _set_merge_progress_visible(self, visible: bool) -> None:
        if visible:
            self.merge_progress_bar["value"] = 0
            self.merge_progress_label["text"] = ""
            self.merge_progress_area.pack(fill=tk.X, pady=(0, 5))
        else:
            self.merge_progress_area.pack_forget()

    def _set_merge_progress(self, pct: int) -> None:
        self.merge_progress_bar["value"] = pct
        self.merge_progress_label["text"] = f"处理中... {pct}%"

    def _set_merge_busy(self, busy: bool) -> None:
        self._merge_busy = busy

    # =========================================================================
    # 拆分功能
    # =========================================================================

    def _setup_split_tab(self) -> None:
        frame = self.split_frame

        # 左右分栏容器
        left_frame = ttk.Frame(frame, width=380)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.pack_propagate(False)

        right_frame = ttk.Frame(frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # === 左侧：文件选择与设置 ===
        # 标题区域
        title_frame = ttk.Frame(left_frame)
        title_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            title_frame, text="拆分 PDF 文件",
            font=("微软雅黑", 12, "bold"), bootstyle="primary"
        ).pack()
        ttk.Label(
            title_frame, text="按页码范围或固定页数拆分",
            font=("微软雅黑", 9), bootstyle="secondary"
        ).pack()

        # 拖拽区域
        drop_container = ttk.LabelFrame(left_frame, text="选择文件")  # type: ignore[attr-defined]
        drop_container.pack(fill=tk.X, pady=5)

        self.drop_frame = ttk.Frame(drop_container, padding=10, relief="groove", borderwidth=2)
        self.drop_frame.pack(fill=tk.X, padx=5, pady=5)

        self.drop_label = ttk.Label(
            self.drop_frame,
            text="拖拽 PDF 文件到这里",
            font=("微软雅黑", 9), bootstyle="secondary", anchor="center"
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True)

        self.drop_frame.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        self.drop_frame.dnd_bind("<<Drop>>", self._on_split_drop)  # type: ignore[attr-defined]
        self.drop_frame.dnd_bind("<<DropEnter>>", self._on_split_drag_enter)  # type: ignore[attr-defined]
        self.drop_frame.dnd_bind("<<DropLeave>>", self._on_split_drag_leave)  # type: ignore[attr-defined]

        # 按钮区域
        btn_row = ttk.Frame(drop_container)
        btn_row.pack(pady=5)
        ttk.Button(
            btn_row, text="选择 PDF 文件", bootstyle="info",
            command=self._split_browse_file, width=15
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_row, text="重置", bootstyle="danger",
            command=self._split_reset, width=8
        ).pack(side=tk.LEFT, padx=2)

        self.split_filepath = tk.StringVar()
        ttk.Entry(
            drop_container, textvariable=self.split_filepath,
            state="readonly", font=("微软雅黑", 8)
        ).pack(fill=tk.X, padx=5, pady=(0, 5))

        # 拆分模式
        mode_frame = ttk.LabelFrame(left_frame, text="拆分模式")  # type: ignore[attr-defined]
        mode_frame.pack(fill=tk.X, pady=5)

        self.split_mode = tk.StringVar(value="range")

        ttk.Radiobutton(
            mode_frame, text="按范围拆分（多文件）",
            variable=self.split_mode, value="range", bootstyle="primary"
        ).grid(row=0, column=0, sticky=tk.W, pady=3, padx=5)

        self.range_var = tk.StringVar(value="1-3,5,7-10")
        self.range_entry = ttk.Entry(
            mode_frame, textvariable=self.range_var,
            font=("微软雅黑", 9)
        )
        self.range_entry.grid(row=1, column=0, padx=5, pady=2, sticky=tk.EW)
        self.range_hint = ttk.Label(
            mode_frame, text="格式：1,3-5,7",
            font=("微软雅黑", 8), bootstyle="secondary"
        )
        self.range_hint.grid(row=1, column=1, padx=3)

        ttk.Radiobutton(
            mode_frame, text="按每份页数拆分",
            variable=self.split_mode, value="chunks", bootstyle="primary"
        ).grid(row=2, column=0, sticky=tk.W, pady=3, padx=5)

        self.chunk_frame = ttk.Frame(mode_frame)
        self.chunk_frame.grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.chunk_var = tk.IntVar(value=3)
        ttk.Spinbox(
            self.chunk_frame, from_=1, to=100,
            textvariable=self.chunk_var, width=6
        ).pack(side=tk.LEFT)
        ttk.Label(
            self.chunk_frame, text="页/份",
            font=("微软雅黑", 9)
        ).pack(side=tk.LEFT, padx=5)

        # 新增：提取页面模式
        ttk.Radiobutton(
            mode_frame, text="提取页面（单文件）",
            variable=self.split_mode, value="extract", bootstyle="primary"
        ).grid(row=4, column=0, sticky=tk.W, pady=3, padx=5)
        ttk.Label(
            mode_frame, text="将选中页面合并为一个 PDF",
            font=("微软雅黑", 8), bootstyle="secondary"
        ).grid(row=5, column=0, sticky=tk.W, padx=5)

        mode_frame.columnconfigure(0, weight=1)

        # 输出预览
        preview_frame = ttk.LabelFrame(left_frame, text="输出预览")  # type: ignore[attr-defined]
        preview_frame.pack(fill=tk.X, pady=5)

        self.split_preview = ttk.Treeview(
            preview_frame, columns=("file", "pages"),
            show="headings", height=3)
        self.split_preview.heading("file", text="文件名")
        self.split_preview.heading("pages", text="页数")
        self.split_preview.column("file", width=200)
        self.split_preview.column("pages", width=50, anchor="center")
        self.split_preview.pack(fill=tk.X, padx=5, pady=5)

        # 操作按钮
        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=8)
        ttk.Button(
            left_frame, text="开始拆分", bootstyle="warning",
            command=self._split_pdf, width=20
        ).pack(pady=5)

        # 进度条
        self.split_progress_area = ttk.Frame(left_frame)
        self.split_progress_bar = ttk.Progressbar(
            self.split_progress_area, bootstyle="warning-striped")
        self.split_progress_bar.pack(fill=tk.X)
        self.split_progress_label = ttk.Label(
            self.split_progress_area, text="", bootstyle="secondary", font=("微软雅黑", 8))
        self.split_progress_label.pack(fill=tk.X)
        self.split_progress_area.pack_forget()

        # === 右侧：PDF 预览 / 缩略图网格 ===
        # 普通预览面板
        self.split_preview_panel = _PDFPreviewPanel(right_frame, self, max_display_width=450, dpi=100)
        self.split_preview_panel.frame.pack(fill=tk.BOTH, expand=True)

        # 缩略图网格（提取模式用）
        self.split_thumbnail_grid = _PageThumbnailGrid(
            right_frame, self,
            on_selection_change=self._on_extract_selection_change,
            thumbnail_size=_THUMBNAIL_SIZE, columns=_THUMBNAIL_COLUMNS
        )
        self.split_thumbnail_grid.frame.pack_forget()  # 初始隐藏

        # 状态栏（底部）
        self.split_status = tk.StringVar(value=_status_msg("就绪"))
        ttk.Label(frame, textvariable=self.split_status, bootstyle="secondary").pack(
            side=tk.BOTTOM, fill=tk.X
        )

        # 实时更新预览
        self.split_mode.trace_add("write", lambda *_: self._update_split_preview())
        self.split_mode.trace_add("write", lambda *_: self._on_split_mode_change())
        self.range_var.trace_add("write", lambda *_: self._update_split_preview())
        self.chunk_var.trace_add("write", lambda *_: self._update_split_preview())
        self.split_filepath.trace_add("write", lambda *_: self._update_split_preview())
        self.split_filepath.trace_add("write", lambda *_: self._on_split_filepath_change_for_preview())

    def _on_split_mode_change(self) -> None:
        """拆分模式切换时更新 UI"""
        mode = self.split_mode.get()
        if mode == "extract":
            # 显示缩略图网格
            self.split_preview_panel.frame.pack_forget()
            self.split_thumbnail_grid.frame.pack(fill=tk.BOTH, expand=True)
            # 加载文件到缩略图网格
            fp = self.split_filepath.get().strip()
            if fp and os.path.isfile(fp):
                self.split_thumbnail_grid.set_file(fp)
        else:
            # 显示普通预览
            self.split_thumbnail_grid.frame.pack_forget()
            self.split_preview_panel.frame.pack(fill=tk.BOTH, expand=True)
            fp = self.split_filepath.get().strip()
            if fp and os.path.isfile(fp):
                self.split_preview_panel.set_file(fp)

    def _on_extract_selection_change(self, selected_pages: set[int]) -> None:
        """提取模式页面选择变化回调"""
        self._update_split_preview()

    def _on_split_filepath_change_for_preview(self) -> None:
        fp = self.split_filepath.get().strip()
        if fp and os.path.isfile(fp):
            mode = self.split_mode.get()
            if mode == "extract":
                self.split_thumbnail_grid.set_file(fp)
            else:
                self.split_preview_panel.set_file(fp)
        else:
            self.split_preview_panel.clear()
            self.split_thumbnail_grid.clear()

    def _update_split_preview(self) -> None:
        """实时更新拆分预览"""
        self.split_preview.delete(*self.split_preview.get_children())
        fp = self.split_filepath.get().strip()
        if not fp or not os.path.isfile(fp):
            return
        try:
            reader = PdfReader(fp)
            total = len(reader.pages)
            reader.stream.close()
        except Exception:
            return

        base = os.path.splitext(os.path.basename(fp))[0]
        mode = self.split_mode.get()

        if mode == "range":
            range_str = self.range_var.get().strip()
            if range_str:
                parts = [p.strip() for p in range_str.split(",") if p.strip()]
                for idx, part in enumerate(parts, 1):
                    pages = self._parse_single_range(part, total)
                    if pages:
                        self.split_preview.insert(
                            "", tk.END, values=(f"{base}_{idx}.pdf", len(pages)))
        elif mode == "chunks":
            cs = self.chunk_var.get()
            if cs < 1:
                return
            nparts = (total + cs - 1) // cs
            for p in range(1, nparts + 1):
                actual = min(cs, total - (p - 1) * cs)
                self.split_preview.insert(
                    "", tk.END, values=(f"{base}_part_{p}.pdf", actual))
        elif mode == "extract":
            # 提取模式：显示选中页面
            selected = self.split_thumbnail_grid.get_selected_pages()
            if selected:
                self.split_preview.insert(
                    "", tk.END, values=(f"{base}_extracted.pdf", len(selected)))

    def _on_split_drag_enter(self, _event) -> None:
        self.drop_frame.configure(bootstyle="success")
        self.drop_label.configure(bootstyle="success")

    def _on_split_drag_leave(self, _event) -> None:
        self.drop_frame.configure(bootstyle="default")
        self.drop_label.configure(bootstyle="secondary")

    def _on_split_drop(self, event) -> None:
        self._on_split_drag_leave(event)
        files = self.root.tk.splitlist(event.data)
        pdfs = [f for f in files if os.path.isfile(f) and f.lower().endswith(SUPPORTED_FORMAT)]
        if pdfs:
            fp = os.path.abspath(pdfs[0])
            self.split_filepath.set(fp)
            self.split_status.set(_status_msg(f"已选择：{os.path.basename(fp)}"))
            self._update_split_preview()
        else:
            self.split_status.set(_status_msg("无效的文件（仅支持 PDF）"))

    def _split_browse_file(self) -> None:
        fp = filedialog.askopenfilename(
            title="选择要拆分的 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if fp:
            fp = os.path.abspath(fp)
            self.split_filepath.set(fp)
            self.split_status.set(_status_msg(f"已选择：{os.path.basename(fp)}"))
            self._update_split_preview()

    def _split_reset(self) -> None:
        """重置拆分页面为初始状态"""
        self.split_filepath.set("")
        self.split_mode.set("range")
        self.range_var.set("1-3,5,7-10")
        self.chunk_var.set(3)
        self.split_preview.delete(*self.split_preview.get_children())
        self.split_preview_panel.clear()
        self.split_thumbnail_grid.clear()
        self.split_status.set(_status_msg("就绪"))

    def _parse_page_ranges(self, page_str: str, total_pages: int) -> Optional[list[int]]:
        page_str = re.sub(r"\s+", "", page_str)
        if not page_str:
            return None

        result: set[int] = set()
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
        fp = self.split_filepath.get().strip()
        if not fp or not os.path.isfile(fp):
            messagebox.showwarning("警告", "请先选择有效的 PDF 文件！")
            return
        if self._split_busy:
            messagebox.showinfo("提示", "正在拆分中，请稍候...")
            return

        mode = self.split_mode.get()

        # 提取模式验证
        if mode == "extract":
            selected = self.split_thumbnail_grid.get_selected_pages()
            if not selected:
                messagebox.showwarning("警告", "请先选择要提取的页面！")
                return

        try:
            reader = PdfReader(fp)
            total_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 PDF：\n{e}")
            return

        base_name = os.path.splitext(os.path.basename(fp))[0]
        output_dir = filedialog.askdirectory(title="选择保存文件夹")
        if not output_dir:
            return

        self._split_busy = True
        self._dispatch(lambda: self.split_status.set(_status_msg(f"{total_pages} 页，正在处理...")))
        thread = threading.Thread(
            target=self._split_worker,
            args=(reader, total_pages, base_name, output_dir, mode),
            daemon=True
        )
        thread.start()

    def _split_worker(self, reader, total_pages, base_name, output_dir, mode):
        try:
            self._dispatch(self._set_split_progress_visible, True)
            if mode == "range":
                self._split_by_range_worker(reader, total_pages, base_name, output_dir)
            elif mode == "chunks":
                self._split_by_chunks_worker(reader, total_pages, base_name, output_dir)
            elif mode == "extract":
                self._split_extract_worker(reader, base_name, output_dir)
        except Exception as e:
            self._dispatch(lambda err=str(e): self._split_fail(err))
        finally:
            self._dispatch(lambda: self._set_split_busy(False))
            self._dispatch(self._set_split_progress_visible, False)

    def _parse_single_range(self, part: str, total_pages: int) -> Optional[list[int]]:
        part = part.strip()
        if not part:
            return None
        if "-" in part:
            try:
                start_str, end_str = part.split("-", 1)
                start, end = int(start_str), int(end_str)
                if start > end:
                    start, end = end, start
                start = max(1, start)
                end = min(total_pages, end)
                result = list(range(start, end + 1))
                return result if result else None
            except (ValueError, IndexError):
                return None
        else:
            try:
                page = int(part)
                if 1 <= page <= total_pages:
                    return [page]
                return None
            except ValueError:
                return None

    def _split_by_range_worker(self, reader, total_pages, base_name, output_dir):
        page_input = self.range_var.get().strip()
        if not page_input:
            self._dispatch(lambda: self._split_fail("请输入页码范围！"))
            return

        parts = [p.strip() for p in page_input.split(",") if p.strip()]
        if not parts:
            self._dispatch(lambda: self._split_fail("页码格式不正确！示例：1,3-5,7"))
            return

        file_count = 0
        for idx, part in enumerate(parts, 1):
            page_numbers = self._parse_single_range(part, total_pages)
            if not page_numbers:
                continue
            writer = PdfWriter()
            for pn in page_numbers:
                writer.add_page(reader.pages[pn - 1])
            out_file = os.path.join(output_dir, f"{base_name}_{idx}.pdf")
            with open(out_file, "wb") as f:
                writer.write(f)
            file_count += 1
            pct = int(idx / len(parts) * 100)
            self._dispatch(lambda p=pct: self._set_split_progress(p))

        if file_count == 0:
            self._dispatch(lambda: self._split_fail("没有有效的页码范围"))
        else:
            self._dispatch(lambda c=file_count, d=output_dir: self._split_range_success(c, d))

    def _split_by_chunks_worker(self, reader, total_pages, base_name, output_dir):
        chunk_size = self.chunk_var.get()
        if chunk_size < 1:
            self._dispatch(lambda: self._split_fail("每份页数必须 >= 1"))
            return

        total_parts = (total_pages + chunk_size - 1) // chunk_size
        part = 1
        for start in range(0, total_pages, chunk_size):
            end = min(start + chunk_size, total_pages)
            writer = PdfWriter()
            for i in range(start, end):
                writer.add_page(reader.pages[i])

            out_file = os.path.join(output_dir, f"{base_name}_part_{part}.pdf")
            with open(out_file, "wb") as f:
                writer.write(f)

            pct = int(part / total_parts * 100)
            self._dispatch(lambda p=pct: self._set_split_progress(p))
            c, t = part, total_parts
            self._dispatch(lambda: self.split_status.set(_status_msg(f"已生成 {c}/{t} 个文件")))
            part += 1

        self._dispatch(lambda c=part - 1, d=output_dir: self._split_chunks_success(c, d))

    def _split_range_success(self, count: int, output_dir: str) -> None:
        self.split_status.set(_status_msg(f"拆分完成（{count} 个文件）"))
        messagebox.showinfo("成功", f"已拆分为 {count} 个文件，保存在：\n{output_dir}")

    def _split_chunks_success(self, count: int, output_dir: str) -> None:
        self.split_status.set(_status_msg(f"拆分完成（{count} 个文件）"))
        messagebox.showinfo("成功", f"已拆分为 {count} 个文件，保存在：\n{output_dir}")

    def _split_extract_worker(self, reader, base_name: str, output_dir: str):
        """提取模式：将选中页面合并为一个文件"""
        page_numbers = self.split_thumbnail_grid.get_selected_pages()
        if not page_numbers:
            self._dispatch(lambda: self._split_fail("请先选择要提取的页面"))
            return

        writer = PdfWriter()
        total = len(page_numbers)
        for idx, pn in enumerate(page_numbers, 1):
            writer.add_page(reader.pages[pn - 1])
            pct = int(idx / total * 100)
            self._dispatch(lambda p=pct: self._set_split_progress(p))

        output_path = os.path.join(output_dir, f"{base_name}_extracted.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

        self._dispatch(lambda p=output_path: self._split_extract_success(p))

    def _split_extract_success(self, output_path: str) -> None:
        self.split_status.set(_status_msg("提取完成"))
        messagebox.showinfo("成功", f"已提取选中页面，保存在：\n{output_path}")

    def _split_fail(self, error: str) -> None:
        self.split_status.set(_status_msg("拆分失败"))
        messagebox.showerror("错误", f"拆分失败：\n{error}")

    def _set_split_progress_visible(self, visible: bool) -> None:
        if visible:
            self.split_progress_bar["value"] = 0
            self.split_progress_label["text"] = ""
            self.split_progress_area.pack(fill=tk.X, pady=(0, 5))
        else:
            self.split_progress_area.pack_forget()

    def _set_split_progress(self, pct: int) -> None:
        self.split_progress_bar["value"] = pct
        self.split_progress_label["text"] = f"处理中... {pct}%"

    def _set_split_busy(self, busy: bool) -> None:
        self._split_busy = busy

    # =========================================================================
    # 旋转页面功能
    # =========================================================================

    def _setup_rotate_tab(self) -> None:
        frame = self.rotate_frame

        # 左右分栏容器
        left_frame = ttk.Frame(frame, width=380)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.pack_propagate(False)

        right_frame = ttk.Frame(frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # === 左侧：文件选择与设置 ===
        # 标题区域
        title_frame = ttk.Frame(left_frame)
        title_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            title_frame, text="旋转 PDF 页面",
            font=("微软雅黑", 12, "bold"), bootstyle="primary"
        ).pack()
        ttk.Label(
            title_frame, text="选择页面范围，设置旋转角度",
            font=("微软雅黑", 9), bootstyle="secondary"
        ).pack()

        # 拖拽区域
        drop_container = ttk.LabelFrame(left_frame, text="选择文件")  # type: ignore[attr-defined]
        drop_container.pack(fill=tk.X, pady=5)

        self.rot_drop_frame = ttk.Frame(drop_container, padding=10, relief="groove", borderwidth=2)
        self.rot_drop_frame.pack(fill=tk.X, padx=5, pady=5)
        self.rot_drop_label = ttk.Label(
            self.rot_drop_frame,
            text="拖拽 PDF 文件到这里",
            font=("微软雅黑", 9), bootstyle="secondary", anchor="center"
        )
        self.rot_drop_label.pack(fill=tk.BOTH, expand=True)

        self.rot_drop_frame.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        self.rot_drop_frame.dnd_bind("<<Drop>>", self._on_rotate_drop)  # type: ignore[attr-defined]
        self.rot_drop_frame.dnd_bind("<<DropEnter>>", self._on_rotate_enter)  # type: ignore[attr-defined]
        self.rot_drop_frame.dnd_bind("<<DropLeave>>", self._on_rotate_leave)  # type: ignore[attr-defined]

        # 按钮区域
        btn_row = ttk.Frame(drop_container)
        btn_row.pack(pady=5)
        ttk.Button(
            btn_row, text="选择 PDF 文件", bootstyle="info",
            command=self._rotate_browse, width=15
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_row, text="重置", bootstyle="danger",
            command=self._rotate_reset, width=8
        ).pack(side=tk.LEFT, padx=2)

        self.rot_filepath = tk.StringVar()
        ttk.Entry(
            drop_container, textvariable=self.rot_filepath,
            state="readonly", font=("微软雅黑", 8)
        ).pack(fill=tk.X, padx=5, pady=(0, 5))

        self.rot_page_info = tk.StringVar(value="")
        ttk.Label(drop_container, textvariable=self.rot_page_info, bootstyle="info").pack(pady=(0, 5))

        # 页面范围
        range_frame = ttk.LabelFrame(left_frame, text="页码范围（留空 = 全部）")  # type: ignore[attr-defined]
        range_frame.pack(fill=tk.X, pady=5)
        self.rot_range_var = tk.StringVar(value="")
        ttk.Entry(
            range_frame, textvariable=self.rot_range_var,
            font=("微软雅黑", 9)
        ).pack(fill=tk.X, padx=5, pady=3, anchor=tk.W)
        ttk.Label(
            range_frame, text="格式：1,3-5,7",
            font=("微软雅黑", 8), bootstyle="secondary"
        ).pack(padx=5, pady=(0, 3))

        # 旋转角度
        angle_frame = ttk.LabelFrame(left_frame, text="旋转角度")  # type: ignore[attr-defined]
        angle_frame.pack(fill=tk.X, pady=5)

        self.rot_angle = tk.IntVar(value=0)
        angles = [
            (0, "不旋转"), (90, "顺时针 90°"),
            (180, "旋转 180°"), (-90, "逆时针 90°"),
        ]
        for i, (val, txt) in enumerate(angles):
            ttk.Radiobutton(
                angle_frame, text=txt, variable=self.rot_angle,
                value=val, bootstyle="primary"
            ).grid(row=i // 2, column=i % 2, sticky=tk.W, padx=15, pady=4)

        # 逐页旋转设置
        per_page_frame = ttk.LabelFrame(left_frame, text="逐页旋转设置")  # type: ignore[attr-defined]
        per_page_frame.pack(fill=tk.X, pady=5)

        self.rotation_tree = ttk.Treeview(
            per_page_frame, columns=("page", "angle"),
            show="headings", height=3
        )
        self.rotation_tree.heading("page", text="页码")
        self.rotation_tree.heading("angle", text="角度")
        self.rotation_tree.column("page", width=80, anchor="center")
        self.rotation_tree.column("angle", width=80, anchor="center")
        self.rotation_tree.pack(fill=tk.X, padx=5, pady=3)

        btn_row = ttk.Frame(per_page_frame)
        btn_row.pack(fill=tk.X, padx=5, pady=3)
        ttk.Button(
            btn_row, text="添加设置", bootstyle="info-outline",
            command=self._add_rotation_setting, width=10
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_row, text="删除选中", bootstyle="secondary-outline",
            command=self._remove_rotation_setting, width=10
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            btn_row, text="清空", bootstyle="secondary-outline",
            command=self._clear_rotation_settings, width=6
        ).pack(side=tk.LEFT, padx=2)

        # 操作按钮
        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=8)
        ttk.Button(
            left_frame, text="旋转并保存", bootstyle="warning",
            command=self._rotate_pages, width=20
        ).pack(pady=5)

        # 进度条
        self.rot_progress_area = ttk.Frame(left_frame)
        self.rot_progress_bar = ttk.Progressbar(
            self.rot_progress_area, bootstyle="danger-striped")
        self.rot_progress_bar.pack(fill=tk.X)
        self.rot_progress_label = ttk.Label(
            self.rot_progress_area, text="", bootstyle="secondary", font=("微软雅黑", 8))
        self.rot_progress_label.pack(fill=tk.X)
        self.rot_progress_area.pack_forget()

        # === 右侧：PDF 预览 ===
        self.rot_preview_panel = _RotationPreviewPanel(right_frame, self, max_display_width=450, dpi=100)
        self.rot_preview_panel.frame.pack(fill=tk.BOTH, expand=True)

        # 状态栏（底部）
        self.rot_status = tk.StringVar(value=_status_msg("就绪"))
        ttk.Label(frame, textvariable=self.rot_status, bootstyle="secondary").pack(
            side=tk.BOTTOM, fill=tk.X
        )

        # 实时旋转预览绑定
        self.rot_angle.trace_add("write", lambda *_: self._update_rotation_preview())

    def _update_rotation_preview(self) -> None:
        """更新旋转预览"""
        angle = self.rot_angle.get()
        self.rot_preview_panel.set_preview_rotation(angle)

    def _on_rotate_enter(self, _event) -> None:
        self.rot_drop_frame.configure(bootstyle="success")
        self.rot_drop_label.configure(bootstyle="success")

    def _on_rotate_leave(self, _event) -> None:
        self.rot_drop_frame.configure(bootstyle="default")
        self.rot_drop_label.configure(bootstyle="secondary")

    def _on_rotate_drop(self, event) -> None:
        self._on_rotate_leave(event)
        files = self.root.tk.splitlist(event.data)
        pdfs = [f for f in files if os.path.isfile(f) and f.lower().endswith(SUPPORTED_FORMAT)]
        if pdfs:
            self._set_rotate_file(os.path.abspath(pdfs[0]))
        else:
            self.rot_status.set(_status_msg("无效的文件（仅支持 PDF）"))

    def _rotate_browse(self) -> None:
        fp = filedialog.askopenfilename(
            title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")]
        )
        if fp:
            self._set_rotate_file(os.path.abspath(fp))

    def _rotate_reset(self) -> None:
        """重置旋转页面为初始状态"""
        self.rot_filepath.set("")
        self.rot_range_var.set("")
        self.rot_angle.set(0)
        self._clear_rotation_settings()
        self.rot_preview_panel.clear()
        self.rot_page_info.set("")
        self.rot_status.set(_status_msg("就绪"))

    def _set_rotate_file(self, fp: str) -> None:
        self.rot_filepath.set(fp)
        self.rot_preview_panel.set_file(fp)
        try:
            r = PdfReader(fp)
            total = len(r.pages)
            r.stream.close()
            self.rot_page_info.set(f"共 {total} 页")
        except Exception:
            self.rot_page_info.set("")
        self.rot_status.set(_status_msg(f"已选择：{os.path.basename(fp)}"))
        # 清空逐页旋转设置
        self._clear_rotation_settings()

    def _add_rotation_setting(self) -> None:
        """添加页面旋转设置"""
        range_str = self.rot_range_var.get().strip()
        angle = self.rot_angle.get()

        if not range_str:
            messagebox.showwarning("提示", "请输入页码范围")
            return

        try:
            r = PdfReader(self.rot_filepath.get())
            total = len(r.pages)
            r.stream.close()
        except Exception:
            messagebox.showerror("错误", "无法读取 PDF 文件")
            return

        pages = self._parse_page_ranges(range_str, total)
        if not pages:
            messagebox.showerror("错误", "页码格式不正确！\n示例：1,3-5,7")
            return

        # 更新 Treeview
        for pn in pages:
            existing = self.rotation_tree.get_children()
            updated = False
            for item in existing:
                values = self.rotation_tree.item(item)["values"]
                if values and values[0] == pn:
                    self.rotation_tree.item(item, values=(pn, f"{angle}°"))
                    updated = True
                    break
            if not updated:
                self.rotation_tree.insert("", tk.END, values=(pn, f"{angle}°"))

        self.rot_status.set(_status_msg(f"已设置 {len(pages)} 页的旋转角度"))

    def _remove_rotation_setting(self) -> None:
        """删除选中的旋转设置"""
        selected = self.rotation_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选中要删除的设置")
            return
        for item in selected:
            self.rotation_tree.delete(item)
        self.rot_status.set(_status_msg("已删除选中设置"))

    def _clear_rotation_settings(self) -> None:
        """清空所有旋转设置"""
        for item in self.rotation_tree.get_children():
            self.rotation_tree.delete(item)

    def _get_rotation_settings(self) -> dict[int, int]:
        """获取所有页面的旋转设置"""
        settings = {}
        for item in self.rotation_tree.get_children():
            values = self.rotation_tree.item(item)["values"]
            if values:
                pn = int(values[0])
                angle_str = str(values[1]).replace("°", "")
                settings[pn] = int(angle_str)
        return settings

    def _rotate_pages(self) -> None:
        fp = self.rot_filepath.get().strip()
        if not fp or not os.path.isfile(fp):
            messagebox.showwarning("警告", "请先选择有效的 PDF 文件！")
            return
        if self._rotate_busy:
            messagebox.showinfo("提示", "正在旋转中，请稍候...")
            return

        try:
            reader = PdfReader(fp)
            total = len(reader.pages)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 PDF：\n{e}")
            return

        # 获取逐页旋转设置
        per_page_settings = self._get_rotation_settings()

        # 如果有逐页设置，使用逐页设置；否则使用默认角度
        if per_page_settings:
            page_numbers = list(per_page_settings.keys())
            angle = None  # 不使用默认角度
        else:
            range_str = self.rot_range_var.get().strip()
            if range_str:
                page_numbers = self._parse_page_ranges(range_str, total)
                if not page_numbers:
                    messagebox.showerror("错误", "页码格式不正确！\n示例：1,3-5,7")
                    return
            else:
                page_numbers = list(range(1, total + 1))
            angle = self.rot_angle.get()

            # 检查是否选择了有效的旋转角度
            if angle == 0:
                messagebox.showwarning("提示", "请选择旋转角度（当前为「不旋转」）")
                return

        output = filedialog.asksaveasfilename(
            title="保存旋转后的文件",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if not output:
            return

        self._rotate_busy = True
        thread = threading.Thread(
            target=self._rotate_worker,
            args=(reader, total, page_numbers, angle, output, per_page_settings),
            daemon=True
        )
        thread.start()

    def _rotate_worker(self, reader, total, page_numbers, angle, output_path, per_page_settings=None):
        try:
            self._dispatch(self._set_rotate_progress_visible, True)
            self._dispatch(lambda: self.rot_status.set(_status_msg("正在旋转...")))

            writer = PdfWriter()
            page_set = set(page_numbers)

            for i, page in enumerate(reader.pages, 1):
                if i in page_set:
                    # 使用逐页设置或默认角度
                    if per_page_settings and i in per_page_settings:
                        page = page.rotate(per_page_settings[i])
                    elif angle is not None:
                        page = page.rotate(angle)
                writer.add_page(page)
                pct = int(i / total * 70)
                self._dispatch(lambda p=pct: self._set_rotate_progress(p))

            self._dispatch(lambda: self._set_rotate_progress(75))
            with open(output_path, "wb") as f:
                writer.write(f)
            self._dispatch(lambda: self._set_rotate_progress(100))
            self._dispatch(lambda: self._rotate_complete(output_path))
        except Exception as e:
            self._dispatch(lambda err=str(e): self._rotate_fail(err))
        finally:
            self._dispatch(lambda: self._set_rotate_busy(False))
            self._dispatch(self._set_rotate_progress_visible, False)

    def _rotate_complete(self, output_path: str) -> None:
        self.rot_status.set(_status_msg("旋转完成！"))
        messagebox.showinfo("成功", f"已保存至：\n{output_path}")

    def _rotate_fail(self, error: str) -> None:
        self.rot_status.set(_status_msg("旋转失败"))
        messagebox.showerror("错误", f"旋转失败：\n{error}")

    def _set_rotate_progress_visible(self, visible: bool) -> None:
        if visible:
            self.rot_progress_bar["value"] = 0
            self.rot_progress_label["text"] = ""
            self.rot_progress_area.pack(fill=tk.X, pady=(0, 5))
        else:
            self.rot_progress_area.pack_forget()

    def _set_rotate_progress(self, pct: int) -> None:
        self.rot_progress_bar["value"] = pct
        self.rot_progress_label["text"] = f"处理中... {pct}%"

    def _set_rotate_busy(self, busy: bool) -> None:
        self._rotate_busy = busy

    # =========================================================================
    # 全局快捷键
    # =========================================================================

    def _on_ctrl_a(self, _event=None) -> None:
        current = self.notebook.select()
        if current == str(self.merge_frame):
            self.merge_listbox.selection_set(0, tk.END)

    def _execute_current_tab(self, _event=None) -> None:
        current = self.notebook.select()
        if current == str(self.merge_frame):
            self._merge_pdfs()
        elif current == str(self.split_frame):
            self._split_pdf()
        elif current == str(self.rotate_frame):
            self._rotate_pages()


# =========================================================================
# PDF 预览面板
# =========================================================================

class _PDFPreviewPanel:
    """分页式 PDF 预览面板，支持按需渲染和缓存。"""

    def __init__(self, parent, app: "PDFToolApp", max_display_width=450, dpi=100):
        self.app = app
        self.dpi = dpi
        self.max_width = max_display_width
        self._filepath = ""
        self._total_pages = 0
        self._current_page = 0
        self._target_page = 0
        self._rendering = False
        self._photo = None
        self._preview_doc = None
        self._build_ui(parent)

    def _build_ui(self, parent) -> None:
        """构建预览 UI 组件。"""
        self.frame = ttk.LabelFrame(parent, text="PDF 预览")  # type: ignore[attr-defined]
        self.frame.pack(padx=5, pady=5)  # 通过外边距实现内边距效果

        # 导航栏
        self.nav = ttk.Frame(self.frame)
        self.nav.pack(fill=tk.X, pady=3)

        self.prev_btn = ttk.Button(
            self.nav, text="◀ 上一页", bootstyle="secondary-outline",
            command=self.go_prev, state=tk.DISABLED, width=10)
        self.prev_btn.pack(side=tk.LEFT, padx=3)

        self.page_label = ttk.Label(
            self.nav, text="请选择 PDF 文件", bootstyle="info",
            font=("微软雅黑", 9, "bold"), anchor="center")
        self.page_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.next_btn = ttk.Button(
            self.nav, text="下一页 ▶", bootstyle="secondary-outline",
            command=self.go_next, state=tk.DISABLED, width=10)
        self.next_btn.pack(side=tk.RIGHT, padx=3)

        # Canvas 显示区域
        canvas_border = ttk.Frame(
            self.frame, bootstyle="default", relief="solid", borderwidth=1)
        canvas_border.pack(fill=tk.BOTH, expand=True, pady=3)

        self.canvas = tk.Canvas(
            canvas_border, bg="#fafafa",
            width=self.max_width, height=350,
            highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 占位符
        self.placeholder = ttk.Label(
            canvas_border, text="选择 PDF 文件后在此处预览",
            bootstyle="secondary", font=("微软雅黑", 10),
            anchor="center")
        self.placeholder.pack(fill=tk.BOTH, expand=True)

        # 加载中提示
        self.loading_label = ttk.Label(
            canvas_border, text="", bootstyle="warning",
            font=("微软雅黑", 9))
        self.loading_label.pack_forget()

        # 初始隐藏导航
        self.nav.pack_forget()

    def set_file(self, filepath: str) -> None:
        """切换当前预览文件。"""
        if not HAS_FITZ:
            self.placeholder.config(text="请安装 PyMuPDF：pip install PyMuPDF")
            return
        if filepath == self._filepath:
            return

        self._close_doc()
        self._filepath = filepath
        self._photo = None

        try:
            doc = fitz.open(filepath)
            self._preview_doc = doc
            self._total_pages = doc.page_count
        except Exception:
            self._total_pages = 0

        if self._total_pages == 0:
            self.placeholder.config(text="无法打开此 PDF 文件")
            self.page_label.config(text="")
            self.prev_btn.config(state=tk.DISABLED)
            self.next_btn.config(state=tk.DISABLED)
            return

        self._current_page = 1
        self.placeholder.pack_forget()
        self.loading_label.pack_forget()
        self.nav.pack(fill=tk.X, pady=3)
        self._render_page(1)

    def go_prev(self) -> None:
        if self._current_page > 1:
            self._render_page(self._current_page - 1)

    def go_next(self) -> None:
        if self._current_page < self._total_pages:
            self._render_page(self._current_page + 1)

    def clear(self) -> None:
        """重置为空状态。"""
        self._close_doc()
        self._filepath = ""
        self._total_pages = 0
        self._current_page = 0
        self._target_page = 0
        self._photo = None
        self.canvas.delete("all")
        self.prev_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.DISABLED)
        self.page_label.config(text="请选择 PDF 文件")
        self.loading_label.pack_forget()
        self.placeholder.config(text="选择 PDF 文件后在此处预览")
        self.placeholder.pack(fill=tk.BOTH, expand=True)
        self.nav.pack_forget()

    def _close_doc(self) -> None:
        if self._preview_doc is not None:
            try:
                self._preview_doc.close()
            except Exception:
                pass
            self._preview_doc = None

    def _render_page(self, page_num: int) -> None:
        """渲染指定页（带缓存检查和后台渲染）。"""
        if page_num < 1 or page_num > self._total_pages:
            return
        if not self._filepath:
            return

        self._target_page = page_num
        cached = _page_cache.get(self._filepath, page_num, self.dpi)
        if cached:
            self._display_image(cached, page_num)
            return

        self._rendering = True
        self.canvas.delete("all")
        self.placeholder.pack_forget()
        self.loading_label.config(text="正在渲染页面...")
        self.loading_label.pack(fill=tk.X, pady=5)
        self.prev_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.DISABLED)
        self.page_label.config(text=f"Loading {page_num} / {self._total_pages}")

        filepath = self._filepath
        dpi = self.dpi
        app = self.app
        panel_ref = self

        def _worker():
            try:
                doc = fitz.open(filepath)
                page = doc[page_num - 1]
                pix = page.get_pixmap(dpi=dpi)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                doc.close()
                app._dispatch(panel_ref._on_render_complete, page_num, img)
            except Exception as e:
                app._dispatch(panel_ref._on_render_error, str(e))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_render_complete(self, page_num: int, image) -> None:
        """渲染完成回调（主线程）。"""
        if self._target_page != page_num:
            return
        self._rendering = False

        _page_cache.set(self._filepath, page_num, image, self.dpi)
        self._display_image(image, page_num)

    def _on_render_error(self, error: str) -> None:
        """渲染错误回调（主线程）。"""
        self._rendering = False
        self.loading_label.pack_forget()
        self.placeholder.config(text=f"渲染失败：{error}")
        self.placeholder.pack(fill=tk.BOTH, expand=True)
        self.prev_btn.config(state=tk.NORMAL)
        self.next_btn.config(state=tk.NORMAL)

    def _display_image(self, image, page_num: int) -> None:
        """在 Canvas 上显示图片（带自适应缩放）。"""
        self.loading_label.pack_forget()

        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 10 or ch < 10:
            cw, ch = self.max_width, 380

        iw, ih = image.size
        scale = min(cw / iw, ch / ih, 1.0)
        if scale < 1.0:
            nw, nh = max(1, int(iw * scale)), max(1, int(ih * scale))
            image = image.resize((nw, nh), Image.Resampling.LANCZOS)

        photo = ImageTk.PhotoImage(image)
        self._photo = photo

        self.canvas.delete("all")
        self.canvas.create_image(cw / 2, ch / 2, image=photo, anchor="center")

        self._current_page = page_num
        self.page_label.config(text=f"Page {page_num} / {self._total_pages}")
        self.prev_btn.config(state=tk.NORMAL if page_num > 1 else tk.DISABLED)
        self.next_btn.config(state=tk.NORMAL if page_num < self._total_pages else tk.DISABLED)


class _RotationPreviewPanel(_PDFPreviewPanel):
    """支持旋转预览的面板"""

    def __init__(self, parent, app: "PDFToolApp", **kwargs):
        super().__init__(parent, app, **kwargs)
        self._preview_rotation = 0

    def set_preview_rotation(self, angle: int) -> None:
        """设置预览旋转角度（实时预览用）"""
        self._preview_rotation = angle
        if self._current_page > 0:
            self._render_page(self._current_page)

    def _display_image(self, image, page_num: int) -> None:
        """重写显示方法，应用旋转"""
        if self._preview_rotation != 0:
            image = image.rotate(-self._preview_rotation, expand=True)
        super()._display_image(image, page_num)


class _PageThumbnailGrid:
    """缩略图网格面板，支持多选页面"""

    def __init__(self, parent, app: "PDFToolApp",
                 on_selection_change: Optional[Callable[[set[int]], None]] = None,
                 thumbnail_size: tuple = _THUMBNAIL_SIZE,
                 columns: int = _THUMBNAIL_COLUMNS):
        self.app = app
        self.on_selection_change = on_selection_change
        self.thumbnail_size = thumbnail_size
        self.columns = columns

        self._filepath = ""
        self._total_pages = 0
        self._selected_pages: set[int] = set()
        self._thumbnails: dict[int, ImageTk.PhotoImage] = {}
        self._thumbnail_frames: dict[int, ttk.Frame] = {}
        self._loading = False
        self._render_queue: list[int] = []
        self._rendering_page = 0

        self._build_ui(parent)

    def _build_ui(self, parent) -> None:
        """构建 UI"""
        self.frame = ttk.LabelFrame(parent, text="页面选择")
        self.frame.pack(padx=5, pady=5)

        # 工具栏
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(fill=tk.X, pady=3)
        ttk.Button(toolbar, text="全选", command=self.select_all,
                   width=6, bootstyle="info-outline").pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="反选", command=self.invert_selection,
                   width=6, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="清空", command=self.clear_selection,
                   width=6, bootstyle="secondary-outline").pack(side=tk.LEFT, padx=2)
        self.selection_label = ttk.Label(toolbar, text="已选: 0 页",
                                         bootstyle="info", font=("微软雅黑", 9))
        self.selection_label.pack(side=tk.RIGHT, padx=5)

        # 滚动区域
        canvas_container = ttk.Frame(self.frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_container, bg="#f5f5f5", highlightthickness=0)
        v_scroll = ttk.Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scroll.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.grid_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")

        # 占位符
        self.placeholder = ttk.Label(
            self.grid_frame, text="选择 PDF 文件后显示页面缩略图",
            bootstyle="secondary", font=("微软雅黑", 10)
        )
        self.placeholder.pack(pady=50)

        # 加载状态
        self.loading_label = ttk.Label(
            self.grid_frame, text="", bootstyle="warning",
            font=("微软雅黑", 9)
        )

        # 绑定事件
        self.grid_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

    def _on_frame_configure(self, _event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def set_file(self, filepath: str) -> None:
        """加载 PDF 并渲染缩略图网格"""
        if not HAS_FITZ:
            self.placeholder.config(text="请安装 PyMuPDF：pip install PyMuPDF")
            return

        self.clear()
        self._filepath = filepath

        try:
            doc = fitz.open(filepath)
            self._total_pages = doc.page_count
            doc.close()
        except Exception:
            self.placeholder.config(text="无法打开此 PDF 文件")
            return

        if self._total_pages == 0:
            self.placeholder.config(text="PDF 文件没有页面")
            return

        self.placeholder.pack_forget()
        self._build_grid()
        self._start_thumbnail_render()

    def _build_grid(self) -> None:
        """构建缩略图网格框架"""
        for i in range(self._total_pages):
            page_num = i + 1
            row, col = i // self.columns, i % self.columns

            frame = ttk.Frame(self.grid_frame, relief="solid", borderwidth=1)
            frame.grid(row=row, column=col, padx=3, pady=3, sticky="nsew")

            # 占位标签
            placeholder = ttk.Label(
                frame, text=f"{page_num}",
                font=("微软雅黑", 8), bootstyle="secondary"
            )
            placeholder.pack(expand=True)

            # 页码标签
            label = ttk.Label(
                frame, text=f"第 {page_num} 页",
                font=("微软雅黑", 7), bootstyle="secondary"
            )
            label.pack()

            frame.bind("<Button-1>", lambda e, p=page_num: self.toggle_page(p))
            placeholder.bind("<Button-1>", lambda e, p=page_num: self.toggle_page(p))

            self._thumbnail_frames[page_num] = frame

    def _start_thumbnail_render(self) -> None:
        """开始后台渲染缩略图"""
        self._render_queue = list(range(1, self._total_pages + 1))
        self._render_next_thumbnail()

    def _render_next_thumbnail(self) -> None:
        """渲染下一个缩略图"""
        if not self._render_queue:
            return

        page_num = self._render_queue.pop(0)
        self._rendering_page = page_num

        filepath = self._filepath
        thumb_size = self.thumbnail_size
        app = self.app
        panel_ref = self

        def _worker():
            try:
                doc = fitz.open(filepath)
                page = doc[page_num - 1]
                zoom = min(thumb_size[0] / page.rect.width, thumb_size[1] / page.rect.height)
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                doc.close()
                app._dispatch(panel_ref._on_thumbnail_ready, page_num, img)
            except Exception:
                app._dispatch(panel_ref._render_next_thumbnail)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_thumbnail_ready(self, page_num: int, image) -> None:
        """缩略图渲染完成"""
        if page_num not in self._thumbnail_frames:
            self._render_next_thumbnail()
            return

        frame = self._thumbnail_frames[page_num]

        # 清除占位符
        for widget in frame.winfo_children():
            widget.destroy()

        # 创建缩略图
        photo = ImageTk.PhotoImage(image)
        self._thumbnails[page_num] = photo

        img_label = ttk.Label(frame, image=photo)
        img_label.pack(expand=True)
        img_label.bind("<Button-1>", lambda e, p=page_num: self.toggle_page(p))

        # 页码标签
        label = ttk.Label(frame, text=f"第 {page_num} 页",
                          font=("微软雅黑", 7), bootstyle="secondary")
        label.pack()

        # 如果已选中，更新样式
        if page_num in self._selected_pages:
            frame.configure(bootstyle="success")

        self._render_next_thumbnail()

    def toggle_page(self, page_num: int) -> None:
        """切换页面选中状态"""
        if page_num not in self._thumbnail_frames:
            return

        frame = self._thumbnail_frames[page_num]

        if page_num in self._selected_pages:
            self._selected_pages.remove(page_num)
            frame.configure(bootstyle="default")
        else:
            self._selected_pages.add(page_num)
            frame.configure(bootstyle="success")

        self._update_selection_label()

        if self.on_selection_change:
            self.on_selection_change(self._selected_pages)

    def select_all(self) -> None:
        """全选所有页面"""
        self._selected_pages = set(range(1, self._total_pages + 1))
        for page_num, frame in self._thumbnail_frames.items():
            frame.configure(bootstyle="success")
        self._update_selection_label()
        if self.on_selection_change:
            self.on_selection_change(self._selected_pages)

    def invert_selection(self) -> None:
        """反选"""
        all_pages = set(range(1, self._total_pages + 1))
        self._selected_pages = all_pages - self._selected_pages
        for page_num, frame in self._thumbnail_frames.items():
            if page_num in self._selected_pages:
                frame.configure(bootstyle="success")
            else:
                frame.configure(bootstyle="default")
        self._update_selection_label()
        if self.on_selection_change:
            self.on_selection_change(self._selected_pages)

    def clear_selection(self) -> None:
        """清空选择"""
        self._selected_pages.clear()
        for frame in self._thumbnail_frames.values():
            frame.configure(bootstyle="default")
        self._update_selection_label()
        if self.on_selection_change:
            self.on_selection_change(self._selected_pages)

    def _update_selection_label(self) -> None:
        """更新选中计数标签"""
        count = len(self._selected_pages)
        self.selection_label.config(text=f"已选: {count} 页")

    def get_selected_pages(self) -> list[int]:
        """获取选中的页码列表（排序后）"""
        return sorted(self._selected_pages)

    def clear(self) -> None:
        """重置为空状态"""
        self._filepath = ""
        self._total_pages = 0
        self._selected_pages.clear()
        self._thumbnails.clear()
        self._render_queue.clear()

        for frame in self._thumbnail_frames.values():
            frame.destroy()
        self._thumbnail_frames.clear()

        self.selection_label.config(text="已选: 0 页")

        # 重新显示占位符
        for widget in self.grid_frame.winfo_children():
            widget.destroy()
        self.placeholder = ttk.Label(
            self.grid_frame, text="选择 PDF 文件后显示页面缩略图",
            bootstyle="secondary", font=("微软雅黑", 10)
        )
        self.placeholder.pack(pady=50)


# =========================================================================
# 辅助函数
# =========================================================================

def _build_display(filepath: str) -> str:
    name = os.path.basename(filepath)
    try:
        r = PdfReader(filepath)
        pages = len(r.pages)
        r.stream.close()
    except Exception:
        pages = 0
    try:
        sz = _format_size(os.path.getsize(filepath))
    except Exception:
        sz = "?"
    return f"{name}  ({pages}页 | {sz})"


def _update_merge_count(app: "PDFToolApp") -> None:
    n = len(app._merge_files)
    if n == 0:
        app.merge_status.set(_status_msg("就绪"))
    else:
        total_pages = sum(app._merge_page_counts.get(f, 0) for f in app._merge_files)
        app.merge_status.set(_status_msg(f"{n} 个文件，共 {total_pages} 页"))


def _current_merge_count(app: "PDFToolApp") -> str:
    n = len(app._merge_files)
    if n == 0:
        return _status_msg("就绪")
    total_pages = sum(app._merge_page_counts.get(f, 0) for f in app._merge_files)
    return _status_msg(f"{n} 个文件，共 {total_pages} 页")


def _update_merge_preview(app: "PDFToolApp") -> None:
    """更新合并预览信息"""
    if not app._merge_files:
        app.merge_preview_info.set("")
        return

    total_pages = sum(app._merge_page_counts.get(f, 0) for f in app._merge_files)
    total_size = 0
    for fp in app._merge_files:
        try:
            total_size += os.path.getsize(fp)
        except Exception:
            pass

    info = f"文件: {len(app._merge_files)} | 页数: {total_pages} | 大小: {_format_size(total_size)}"
    app.merge_preview_info.set(info)


def main():
    root = TkinterDnD.Tk()
    _app = PDFToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
