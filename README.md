# 发票 Word 生成器

BITFSAE 车队运营组内部使用。

作用：从发票 `PDF/XML` 生成报账说明、验收单和中间校对表；不做通用模板引擎，不保证任意 Word 都能直接套用。

## 现在怎么用

1. 准备一个发票文件夹，里面放本批次的发票 PDF、XML、查验单、付款截图。
2. 打开程序，先选：
   - 模板方案
   - 资料档案
3. 确认发票文件夹、输出目录、文档日期、购买方抬头。
4. 普通情况 `OCR` 用“关闭”。
5. 点击“开始生成”。
6. 如果校验失败，打开输出目录里的 `invoice_intermediate.xlsx` 修正后再生成。

## 输入要求

- 同一批次发票必须属于同一购买方抬头。
- 复杂发票优先提供 XML。
- 模板不是任意 `.docx`，必须符合本项目当前表格结构。
- 发票 PDF 推荐命名：`序号+品类+金额+发票.pdf`
- XML 推荐命名：`序号.xml`

## 模板方案

程序现在支持“模板方案”，不是每次重新找两个 Word 文件。

一个模板方案保存这些内容：

- 报账说明模板路径
- 验收单模板路径
- 报账说明输出文件名
- 验收单输出文件名

主界面切换方案后，会自动带出对应模板和输出文件名。手动改了模板路径或文件名后，可以在“设置”里覆盖保存到当前方案，或者另存成新方案。

## 资料档案

资料档案保存在 `settings.json`，适合一个同学代多人处理报账。

每个档案可保存：

- 赛季
- 学号
- 姓名
- 联系方式
- 开户行
- 卡号

主界面负责选择档案；新增、修改、删除在“设置”里完成。生成报账说明时，程序会把所选档案的学号、姓名、联系方式、开户行、卡号写入对应位置。

## OCR

只有在 `auto` 或 `always` 模式下才会调用阿里云 OCR。

如果需要阿里云文字识别 OCR 密钥，可以联系软件作者，或者自己去阿里云申请。

先配置环境变量：

```bash
export ALIBABA_CLOUD_ACCESS_KEY_ID=你的AccessKeyId
export ALIBABA_CLOUD_ACCESS_KEY_SECRET=你的AccessKeySecret
```

含义：

- `off`：只用 XML 和 PDF 文本解析
- `auto`：解析失败或金额不闭合时补救
- `always`：没有 XML 匹配的 PDF 都走 OCR

## 启动

```bash
pip install -r requirements.txt
python3 main.py
```

CLI 也能直接用：

```bash
python3 generate_invoice_docs.py \
  --invoice-dir ../发票 \
  --reimburse-template ../默认报账说明模板.docx \
  --acceptance-template ../默认验收单模板.docx \
  --output-dir output
```

## 设置文件

程序设置保存在：

- macOS：`~/Library/Application Support/InvoiceWordBuilder/settings.json`
- Windows：`%APPDATA%/InvoiceWordBuilder/settings.json`
- Linux：`~/.config/InvoiceWordBuilder/settings.json`
