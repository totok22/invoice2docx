![Pasted image 20260513124709](https://img.p0li.space/img/20260513160008420.png)

https://ocr.console.aliyun.com/overview

https://help.aliyun.com/zh/ocr/developer-reference/api-ocr-api-2021-07-07-recognizeinvoice?spm=a2c4g.11186623.0.i0

文字识别 ocr -> 票据凭证识别 -> 增值税发票识别
有免费额度，之后按量计费，很便宜。



在macOS中添加全局环境变量，**当前用户全局生效**（针对当前用户的所有终端窗口）：
#### 1. 打开或创建配置文件

在终端中输入以下命令，使用 `nano` 编辑器打开配置文件：

```bash
nano ~/.zshrc
```

#### 2. 写入环境变量语法

在文件的末尾添加环境变量。常见的添加方式有两种：
- **添加自定义变量：**
    ```bash
    export MY_VAR="your_value"
    ```

#### 3. 保存并退出编辑器

1. 按 `Ctrl + O` 保存文件。    
2. 按 `回车键`（Enter）确认文件名。
3. 按 `Ctrl + X` 退出编辑器。
#### 4. 使配置立即生效
输入以下命令重新加载配置文件，使其在当前终端窗口立即生效：
```bash
source ~/.zshrc
```
