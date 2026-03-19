# Panda - 企业会计小工具集

个人使用的会计工作辅助工具集合。

## 目录结构

```
panda/
├── tools/                    # 工具脚本
│   ├── invoice/              # 发票相关
│   ├── voucher/              # 凭证处理
│   ├── report/               # 报表工具
│   └── utils/                # 公共函数
├── data/                     # 输入数据（不上传敏感信息）
├── output/                   # 输出结果
├── requirements.txt          # Python 依赖
└── README.md
```

## 使用方法

1. 创建虚拟环境：
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   pip install -r requirements.txt
   ```

2. 将数据文件放入 `data/` 目录

3. 运行工具：
   ```bash
   python tools/invoice/process_invoices.py
   ```

## 注意事项

- `data/` 和 `output/` 目录已加入 `.gitignore`，敏感财务数据不会上传到 GitHub
- 建议定期备份本地数据