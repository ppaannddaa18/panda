"""文件加载组件"""

import tkinter as tk
from tkinter import ttk, filedialog
from typing import Optional, Callable
from pathlib import Path

try:
    from tkinterdnd2 import DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


class FileLoader(ttk.LabelFrame):
    """文件加载组件（支持拖拽）"""

    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        on_file_loaded: Optional[Callable[[str], None]] = None
    ):
        """
        初始化文件加载组件

        Args:
            parent: 父组件
            title: 标题
            on_file_loaded: 文件加载回调
        """
        super().__init__(parent, text=f" {title} ", padding=10)
        self.title = title
        self.on_file_loaded = on_file_loaded
        self.file_path: Optional[str] = None

        self._setup_ui()

    def _setup_ui(self):
        """设置 UI"""
        # 拖拽区域
        self.drop_frame = tk.Frame(
            self,
            bg="#e8f4fd",
            relief="ridge",
            bd=2,
            height=80
        )
        self.drop_frame.pack(fill=tk.X, pady=(0, 5))
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="拖拽文件到此处\n或点击选择文件",
            bg="#e8f4fd",
            fg="#666666",
            font=("Arial", 9)
        )
        self.drop_label.pack(expand=True)

        # 绑定点击事件
        self.drop_frame.bind("<Button-1>", self._on_click)
        self.drop_label.bind("<Button-1>", self._on_click)

        # 绑定拖拽事件
        if HAS_DND:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)

        # 文件信息
        self.file_label = ttk.Label(self, text="未选择文件", foreground="gray")
        self.file_label.pack(fill=tk.X)

    def _on_click(self, event):
        """点击选择文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择{self.title}",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self._load_file(file_path)

    def _on_drop(self, event):
        """拖拽文件"""
        file_path = event.data
        # 处理 Windows 路径格式
        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]

        if file_path.lower().endswith((".xlsx", ".xls")):
            self._load_file(file_path)

    def _load_file(self, file_path: str):
        """加载文件"""
        self.file_path = file_path
        path = Path(file_path)

        self.file_label.config(
            text=f"✓ {path.name}",
            foreground="green"
        )
        self.drop_label.config(
            text=f"已加载:\n{path.name}",
            fg="green"
        )

        if self.on_file_loaded:
            self.on_file_loaded(file_path)

    def clear(self):
        """清除文件"""
        self.file_path = None
        self.file_label.config(text="未选择文件", foreground="gray")
        self.drop_label.config(
            text="拖拽文件到此处\n或点击选择文件",
            fg="#666666"
        )
