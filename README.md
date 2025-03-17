# 自动生成总结工具

领导丢给你一大堆文件让你写总结，限你赶快弄完。还在为写总结而发愁吗？
现在连AI网页都不用打开啦！双击鼠标即可生成————综合总结.docx

本项目提供了一个自动化工具，能够利用 OpenAI 的 API 对 PDF、DOCX、PPTX 和 XLSX 等多种文档格式进行内容提取与总结，并生成一份综合摘要的 Word 文档。

<img src="image.png" width="400px" />

---

## 功能简介

- **多格式支持**：支持PDF、DOCX、PPTX、XLSX格式文件。
- **自动总结**：调用 OpenAI 的接口，自动生成准确、清晰的综合摘要。
- **便捷交互**：提供了两个不同的版本，分别满足不同用户需求。

## 两个版本说明

- **版本一（便捷版）**：API 的URL与密钥固定在代码内，直接运行即可。
- **版本二（灵活版）**：带有图形界面，允许用户随时修改 OpenAI API 的URL与密钥，更加灵活。

---

## 如何使用

### 一、直接运行Python脚本

#### 第一个版本（便捷版）
1. 在代码中提前填写好 OpenAI 的 URL 与 API Key。
2. 确保所需总结的文档与脚本在同一目录。
3. 运行脚本：

```bash
python convenient_version.py
```

#### 第二个版本（灵活版）
1. 首次运行时，程序会自动提示输入 API URL 和密钥，保存至注册表。
2. 确保文档和脚本在同一目录，直接运行即可：

```bash
python flexible_version.py
```

### 二、打包为可直接双击的exe文件（更推荐）

两个版本均可打包为EXE文件，便于直接运行：

- 安装PyInstaller（如未安装）：

```bash
pip install pyinstaller
```

#### 打包命令

##### 版本一（便捷版）

```bash
pyinstaller -F -w --icon=filesummary.ico convenient_version.py
```

##### 版本二（灵活版）

```bash
pyinstaller -F -w --icon=filesummary.ico flexible_version.py
```

打包完成后，可在`dist`目录中找到可执行文件。

---

## GitHub文件说明

本仓库提供以下文件：
- `convenient_version.py`：版本一Python脚本。
- `flexible_version.py`：版本二Python脚本。
- `flexible_version.exe`：已打包好的版本二可执行程序。（可直接下载并使用）
- `filesummary.ico`：程序使用的图标文件。

---

## 注意事项

- 程序需要稳定的网络环境以调用 OpenAI API。
- 请确保OpenAI API Key有效且URL稳定。
- GitHub上有大量免费的OpenAI API，例如
- [free_chatgpt_api](https://github.com/popjane/free_chatgpt_api)
- 免费api等仅供学习参考，与本项目无关。
---

## 作者
HenryYhZhang

---

## 版权声明
本项目遵循 MIT 开源协议。

