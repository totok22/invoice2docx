# 发票 Word 生成器

FSAE 车队发票处理桌面应用。从发票 PDF/XML 自动生成报账说明和验收单 Word 文档。

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 直接运行（开发模式）
python main.py

# 或用 flet 运行
flet run main.py
```

首次启动会自动带出脚本默认路径：

- 发票文件夹：`../发票`
- 报账说明模板：`../第三组报账说明.docx`
- 验收单模板：`../第三组验收单.docx`
- 输出目录：`output/`
- 默认输出文件名：`第四组报账说明.docx`、`第四组验收单.docx`

界面右上角的"设置与说明"可查看当前默认路径、OCR 前提、校验规则和推荐做法。调整路径或参数后，可以点击"保存为默认"；勾选"自动记住本次选择"后，选择文件夹、模板或运行时会自动保存。

设置保存位置：

- macOS：`~/Library/Application Support/InvoiceWordBuilder/settings.json`
- Windows：`%APPDATA%/InvoiceWordBuilder/settings.json`
- Linux：`~/.config/InvoiceWordBuilder/settings.json`

## CLI 工具

桌面应用基于 `generate_invoice_docs.py` 命令行工具构建，也可以直接使用：

```bash
python generate_invoice_docs.py \
  --invoice-dir ../发票 \
  --reimburse-template ../第三组报账说明.docx \
  --acceptance-template ../第三组验收单.docx \
  --output-dir output \
  --storage-location 工训楼 \
  --document-date 2026年5月13日 \
  --expected-buyer-name 北京理工大学教育基金会
```

常用参数：

- `--from-xlsx 路径`：从人工修正后的中间表重新生成 Word
- `--template-xlsx`：只生成一份带示例行的中间表模板
- `--ocr-mode off|auto|always`：阿里云 OCR 模式
- `--allow-risky-generate`：存在严重差异时仍生成 Word
- `--expected-buyer-name`：指定期望购买方抬头
- `--reimburse-name 文件名.docx`、`--acceptance-name 文件名.docx`：指定输出文件名

## 打包为独立应用

macOS:
```bash
./build_mac.sh
```

Windows:
```bat
build_win.bat
```

打包后的应用在 `dist/` 目录下，双击即可运行，无需安装 Python。

## 功能

- 选择发票文件夹，自动解析 PDF + XML
- 校验金额一致性、购买方信息
- 生成报账说明和验收单 Word 文档
- 导出可修正的中间 XLSX 表
- 支持从修正后的中间表重新生成
- 支持阿里云 OCR：关闭 / 自动补救 / 强制 OCR
- 自动记住常用路径、模板、输出文件名和校验参数
- 实时进度显示和问题清单

## 操作流程

1. 把发票 PDF/XML、查验单、付款截图放入同一个发票文件夹。复杂发票、明细较多的发票，推荐附上 XML。
2. 确认发票文件夹、两份模板和输出目录。
3. OCR 默认关闭；普通 XML/PDF 能解析时不需要开启。
4. 点击"开始生成"。
5. 如果校验通过，直接使用输出的 Word。
6. 如果校验失败，打开输出目录中的 `invoice_intermediate.xlsx`，优先修正浅黄色列后，在"从中间表生成"中选择该 XLSX 重新生成。

## 推荐做法

- 复杂发票、明细较多的发票，尽量同时提供 XML，稳定性通常优于 PDF 文本解析和 OCR。
- XML 命名用"序号.xml"即可，例如 `28.xml`。
- 发票 PDF 建议以"序号+品类+金额+发票.pdf"命名，例如 `28+电子元件+390.05+发票.pdf`。
- 查验单、付款截图可以放在同一个发票文件夹中，不会进入 Word 明细。
- 中间表里浅黄色列是主要修正区；隐藏列通常不用改，需要时可在 Excel/WPS 中取消隐藏。

## OCR 前提

阿里云 OCR 只在选择"自动补救"或"强制 OCR"时调用。使用前必须配置环境变量：

```bash
export ALIBABA_CLOUD_ACCESS_KEY_ID=你的AccessKeyId
export ALIBABA_CLOUD_ACCESS_KEY_SECRET=你的AccessKeySecret
```

模式说明：

- 关闭：只用 XML 和 PDF 文本解析。
- 自动补救：PDF 文本解析失败或金额不闭合时才调用 OCR。
- 强制 OCR：没有 XML 匹配的 PDF 都走 OCR。

## 校验规则

- 同一批发票的购买方抬头必须一致。
- 每张发票都会校验"发票价税合计 = 明细含税金额合计"，差 0.01 也会默认停止生成 Word。
- 勾选"即使有校验错误也强制生成 Word"只建议用于临时核对版；审计文件仍会保留风险标记。

## 处理规则

- 数据优先级：人工修正 XLSX > XML > PDF 文本解析 > 阿里云 OCR > 人工修正中间表。
- 有 XML 的立创发票优先读取 XML，并用 PDF 发票号码匹配对应文件。
- 非 XML 发票读取 PDF 中的发票号码、销售方和总额。
- 连续出现、项目名称相同、且没有数量的折扣/调整行，会合并到上一条同名物品行。
- 同一个发票文件夹中的所有发票必须拥有一致的购买方抬头。
- 每张发票都严格校验 `整理后明细含税金额合计 == 发票价税合计`。
- 验收单展开每一条整理后的物品明细。
- 报账说明每张发票只生成一行，取第一行物品名和发票价税合计。

## 依赖

- Python 3.11+
- flet (UI 框架)
- python-docx (Word 文档)
- openpyxl (Excel 读写)
- pypdf (PDF 文本提取)
- alibabacloud-ocr-api20210707 等阿里云 OCR SDK（仅 OCR 模式需要）
