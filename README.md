# Panda - 企业会计小工具集

个人使用的会计工作辅助工具集合，包含发票处理、PDF 操作、Excel 搜索等实用工具。

## 功能概览

| 工具 | 功能描述 |
|------|----------|
| PDF 工具箱 | PDF 合并、拆分、旋转、预览 |
| PDF OCR 识别 | 批量识别 PDF 中的文字，支持框选区域 |
| Excel 多值搜索 | 多条件组合搜索 Excel 数据 |
| 发票处理 | 批量处理发票数据，金额汇总统计 |

## 工具详情

### 1. PDF 工具箱 (`tools/pdf/`)

功能特性：
- 合并多个 PDF 文件
- 按页码范围/固定页数拆分 PDF
- 旋转 PDF 页面
- 拖拽操作支持
- 缩略图预览

启动方式：
```bash
python tools/pdf/pdf_tool.py
```

### 2. PDF OCR 识别 (`tools/PDFOCR/`)

功能特性：
- 基于 RapidOCR (PP-OCRv4) 中文识别
- 可视化框选识别区域
- 模板系统（保存/加载配置）
- 图像预处理（旋转、亮度、对比度）
- 导出为 Excel/CSV

启动方式：
```bash
python tools/PDFOCR/main.py
```

### 3. Excel 多值搜索 (`tools/excel/`)

功能特性：
- 2-3 个值的组合搜索
- 支持包含/精确匹配模式
- 支持搜索单元格批注
- 拖拽加载文件

启动方式：
```bash
python tools/excel/multi_value_search.py
```

### 4. 发票处理 (`tools/invoice/`)

功能特性：
- 批量读取发票数据
- 金额汇总统计
- 数字转中文大写金额

启动方式：
```bash
python tools/invoice/process_invoices.py
```

## 目录结构

```
panda/
├── tools/
│   ├── pdf/              # PDF 工具箱
│   ├── PDFOCR/           # PDF OCR 识别
│   ├── excel/            # Excel 多值搜索
│   ├── invoice/          # 发票处理
│   ├── report/           # 报表工具（待开发）
│   ├── voucher/          # 凭证处理（待开发）
│   └── utils/            # 公共函数库
├── data/                 # 输入数据
├── output/               # 输出结果
├── dist/                 # 打包输出
├── requirements.txt
└── README.md
```

## 安装

1. 克隆仓库：
```bash
git clone https://github.com/ppaannddaa18/panda.git
cd panda
```

2. 创建虚拟环境：
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # Linux/Mac
```

3. 安装依赖：
```bash
pip install -r requirements.txt
```

## 注意事项

- `data/` 和 `output/` 目录已加入 `.gitignore`，敏感财务数据不会上传
- PDF OCR 工具需要单独安装依赖：`pip install -r tools/PDFOCR/requirements.txt`

## License

MIT
