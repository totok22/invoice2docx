# 发票 Word 生成器

BITFSAE 车队运营组内部使用。

- 版本：V0.1
- 日期：2026-5-13
- GitHub 仓库：https://github.com/totok22/invoice2docx

## 软件用途

本工具用于把一批发票文件生成：

- `报账说明.docx`
- `验收单.docx`
- `invoice_intermediate.xlsx` 中间校对表
- 审计用 CSV/JSON 文件

它不是通用 Word 模板引擎。默认模板已经按当前报账说明和验收单格式整理过，换模板时必须保持相同的表格结构。

## 推荐文件准备方式

1. 新建一个本次报账专用文件夹。
2. 把本批次的发票 PDF、XML、查验单、付款截图等材料放进去。
3. 复杂发票、条目很多的发票，尽量同时放 XML 文件。
4. 同一批次发票应属于同一购买方抬头。

推荐命名：

- 发票 PDF：`序号+品类+金额+发票.pdf`
- 发票 XML：`序号.xml`

例如：

- `1+电子元件+128.50+发票.pdf`
- `1.xml`

程序会优先用 XML 解析发票明细；没有 XML 时，再尝试从 PDF 文本解析。付款截图、查验单等辅助材料不会写入导出的 Word。

## 图形界面使用步骤

1. 打开程序。
2. 在主界面选择“模板方案”。
3. 在主界面选择“资料档案”。
4. 选择“发票文件夹”。
5. 确认“报账说明模板”和“验收单模板”路径。
6. 选择“输出目录”。
7. 填好“文档日期”“存放地点”“期望购买方抬头”。
8. 普通情况把“阿里云 OCR”设为“关闭”。
9. 点击“开始生成”。

生成结果会放在输出目录下的新子文件夹里。每次运行都会新建唯一子文件夹，不会直接覆盖上一次结果。

## 校验失败时怎么处理

如果程序提示校验失败，通常是发票金额、明细合计、购买方抬头或文件命名存在问题。

处理步骤：

1. 打开本次输出目录。
2. 找到 `invoice_intermediate.xlsx`。
3. 在表格里修正识别错误、物资名称、数量、单价、金额或存放地点。
4. 回到程序主界面。
5. 在“从中间表生成（可选）”里选择修正后的 `.xlsx`。
6. 再点“开始生成”。

如果确认校验问题不影响本次报账，也可以勾选“允许带风险生成 Word”。这个选项只适合人工已经核对过的情况。

## 模板方案

模板方案用于保存一组常用模板和输出文件名。

一个模板方案包含：

- 报账说明模板路径
- 验收单模板路径
- 报账说明输出文件名
- 验收单输出文件名

主界面切换方案后，会自动带出对应模板和输出文件名。手动改了模板路径或文件名后，可以在“设置”里覆盖保存到当前方案，或者按名称另存为新方案。

## 资料档案

资料档案保存在程序设置文件里，适合一个同学代多人处理报账。

每个档案可保存：

- 赛季
- 学号
- 姓名
- 联系方式
- 开户行
- 卡号

主界面负责选择档案；新增、修改、删除在“设置”里完成。生成报账说明时，程序会把所选档案的学号、姓名、联系方式、开户行、卡号写入对应位置。

## OCR 模式

主界面“阿里云 OCR”下拉框有三个选项：

- “关闭”：不调用 OCR，只使用 XML 和 PDF 文本解析。普通情况选这个。
- “自动补救”：解析失败或金额不闭合时，才尝试调用 OCR。
- “强制 OCR”：没有匹配 XML 的 PDF 都会调用 OCR。

只有选择“自动补救”或“强制 OCR”时，才需要配置阿里云 OCR 密钥。

## 阿里云 OCR 密钥申请

阿里云入口：

- OCR 控制台：https://ocr.console.aliyun.com/overview
- 增值税发票识别 API 文档：https://help.aliyun.com/zh/ocr/developer-reference/api-ocr-api-2021-07-07-recognizeinvoice

大致步骤：

1. 登录阿里云账号。
2. 进入“文字识别 OCR”。
3. 找到“票据凭证识别”。
4. 开通“增值税发票识别”相关能力。
5. 在阿里云控制台创建 AccessKey。
6. 记录 `AccessKeyId` 和 `AccessKeySecret`。
7. 按下面的方法配置到本机环境变量。

注意：阿里云一般会有免费额度，超出后按量计费。具体价格和额度以阿里云当前页面为准。

## 配置 OCR 环境变量

程序读取这两个环境变量：

- `ALIBABA_CLOUD_ACCESS_KEY_ID`
- `ALIBABA_CLOUD_ACCESS_KEY_SECRET`

macOS 或 Linux 终端临时配置：

```bash
export ALIBABA_CLOUD_ACCESS_KEY_ID="你的AccessKeyId"
export ALIBABA_CLOUD_ACCESS_KEY_SECRET="你的AccessKeySecret"
python3 main.py
```

macOS 当前用户长期配置：

```bash
nano ~/.zshrc
```

在文件末尾加入：

```bash
export ALIBABA_CLOUD_ACCESS_KEY_ID="你的AccessKeyId"
export ALIBABA_CLOUD_ACCESS_KEY_SECRET="你的AccessKeySecret"
```

保存后执行：

```bash
source ~/.zshrc
```

Windows PowerShell 当前用户长期配置：

```powershell
[Environment]::SetEnvironmentVariable("ALIBABA_CLOUD_ACCESS_KEY_ID", "你的AccessKeyId", "User")
[Environment]::SetEnvironmentVariable("ALIBABA_CLOUD_ACCESS_KEY_SECRET", "你的AccessKeySecret", "User")
```

配置后重新打开程序。程序“说明”里会提示是否检测到 OCR 环境变量。

## 从源码启动

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python main.py
```

Windows 可按实际 Python 路径调整命令，例如：

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python main.py
```

## 命令行生成

图形界面是主要入口。需要脚本化时，也可以使用 CLI：

```bash
python3 generate_invoice_docs.py \
  --invoice-dir ./invoices \
  --reimburse-template ./默认报账说明模板.docx \
  --acceptance-template ./默认验收单模板.docx \
  --output-dir ./output
```

常用参数：

- `--invoice-dir`：发票文件夹。
- `--output-dir`：输出目录。
- `--document-date`：写入 Word 的日期。
- `--storage-location`：验收单存放地点。
- `--from-xlsx`：从修正后的中间表生成 Word。
- `--ocr-mode off|auto|always`：命令行内部值，对应界面的“关闭/自动补救/强制 OCR”。

## 设置文件位置

程序设置保存在：

- macOS：`~/Library/Application Support/InvoiceWordBuilder/settings.json`
- Windows：`%APPDATA%/InvoiceWordBuilder/settings.json`
- Linux：`~/.config/InvoiceWordBuilder/settings.json`

如果设置混乱，可以关闭程序后删除这个文件，再重新打开程序。
