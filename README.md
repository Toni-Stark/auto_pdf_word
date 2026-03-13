# PDF 自动转换工具

监控 PDF 文件夹，自动调用 WPS 完成 PDF 转 Word，转换完成后清空 PDF 文件夹。

## 环境要求

- Python 3.8+
- WPS Office（已安装 PDF 转换功能）

安装依赖：

```bash
pip install uiautomation pillow easyocr pyautogui numpy requests torch
```

## 配置

编辑 `config.txt`：

```
PDF_FOLDER=F:\pdf_word\pdf
WORD_FOLDER=F:\pdf_word\word
WXWORK_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=你的key
CHECK_INTERVAL=5
CAPTURE_INTERVAL=60
```

编辑 `pdf2word.ps1`，确认 WPS 安装路径正确：

```powershell
$WPS_HOME = 'F:\wps\WPS Office\12.1.0.25225\office6'
```

将 `WPS PDF转换.lnk`，wps pdf转换 快捷方式放到项目根目录


## 使用

将 PDF 文件放入 `pdf` 文件夹，然后运行：

```bash
python all_step.py
```

转换后的 Word 文件在 `word` 文件夹中。

## 注意

- 首次运行会下载 OCR 模型，需要等待
- 运行期间不要最小化 WPS 窗口
- 如果 WPS 需要登录，程序会截屏二维码发送到企业微信，扫码后自动继续
