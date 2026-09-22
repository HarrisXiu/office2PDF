# Office2PDF v5.x（Office 全能 PDF 工具）简体中文说明

> [English](./README.md) | [日本語](./readme.ja.md) | 简体中文 | [繁體中文](./readme.zh-tw.md)

> 为需要可靠、离线批量 PDF 处理的办公人士设计。

**一款面向办公场景的批量 PDF 转换与处理工具。**
批量将 Office 文件和图片转换为 PDF，并可添加水印、页码、拆分/合并，以及灵活的命名规则，几步点击即可完成。

> **离线转换：** 本应用通过 COM 自动化调用 Microsoft Office 或 WPS Office。
> Office 文档转换需要安装对应办公组件并可通过 COM 调用；PDF、图片处理无需 Office/WPS。所有转换均在本地完成。

### 窗口缩放与转换引擎

- 开始转换、取消和进度条固定在底部；设置区支持滚动，日志可折叠，文件列表与设置区之间可拖动调节高度。
- 窄窗口下设置分组上下排列；较宽的设置行可用底部横向滚动条查看。
- 转换引擎可选择「自动（优先 Office）」「Microsoft Office」「WPS Office」，随配置及预设保存。
- 自动模式仅在对应 Office 组件无法创建时尝试 WPS；文件打开或导出失败会记录错误，不会静默更换引擎。
- 引擎同时用于 Word/Excel/PPT 转换、水印预览与工作表读取。当前支持原有 Office 文件格式，不新增 `.wps/.et/.dps` 原生格式。
- WPS 的支持范围取决于所安装版本的 COM 能力；安装了 WPS 但未注册可用自动化组件时，日志会列出失败的组件名称。
- 新版可执行文件：`dist/Office2PDF.exe`。

---

## 🌟 主要功能（v5.x）

### 批量转换
- 支持以下文件类型转换为 PDF：
  - **Word**：`.doc`、`.docx`
  - **Excel**：`.xls`、`.xlsx`、`.xlsm`
  - **PowerPoint**：`.ppt`、`.pptx`
  - **图片**：`.jpg`、`.jpeg`、`.png`
  - **PDF**：重新保存/规范化处理

### 灵活的拆分 / 合并
- **全部合并**为单个 PDF
- **按页拆分**（Word / PowerPoint / PDF / Excel）
- **按工作表拆分**（Excel）

### 范围选择
- 双击列表项指定目标页码范围（Word / PowerPoint / PDF）：
  - 示例：`1-3,5,8-` / `2` / `-3` / `1-`
- Excel：选择一个或多个**工作表名称**

### 水印与页码
- 最多 **2 个水印**，可设置位置、不透明度、字体、大小和颜色
- **页码**支持格式模板（如 `- {n} / {total} -`）
  - *（当前版本）* 页码以**黑色、固定 10.5pt** 渲染
- **预览**：批量执行前可用**第一页**生成预览 PDF

### 命名规则与预设
- 使用标签的灵活输出命名模板（见下文）
- 保存/读取预设（JSON 配置）

### 安全与优化
- 批量设置 PDF 密码
- 移除元数据
- 压缩 PDF（取决于 pypdf 版本支持）

### 易用性
- 拖拽添加文件
- 列表排序、移除、清空
- 输出选项：
  - 与源文件相同文件夹，或自定义文件夹
- 后处理操作：
  - 打开 PDF / 打开输出文件夹 / 清空列表

### 🌐 多语言界面
- 支持日语 / 英语 / 简体中文 / 繁体中文
- 首次启动按系统语言自动选择，也可通过菜单栏 **Language** 手动切换（保存后重启生效）

---

## 🏷️ 可用标签

标签可用于**命名规则**和**水印文本**，处理时自动替换。

- `{name}`：原始文件名（不含扩展名）
- `{sheet}`：Excel 工作表名称（非 Excel 为空）
- `{parent}`：上级文件夹名
- `{seq}`：全局序号
- `{fseq}`：单个输入文件内的序号
- `{pseq}`：页面序号（拆分/统计页数时）
- `{total}`：输入文件总数
- `{ptotal}`：当前文件/单元的总页数
- `{username}`：Windows 登录用户名
- `{rand}`：4 位随机数字（如 `0427`）

### 日期/时间标签
- `{date:yyyy-mm-dd}` → `2026-02-06`
- `{date:yyyy/mm/dd HH:MM}` → `2026/02/06 16:54`

支持的占位符：
- `yyyy` 年 / `mm` 月 / `dd` 日
- `HH` 时 / `MM` 分 / `SS` 秒

---

## 🚀 使用方法

### 系统要求
- Windows 10 / 11（64 位）
- **已安装 Microsoft Office 或支持 COM 的 WPS Office**（所需的文字 / 表格 / 演示组件）

### 基本步骤
1. 启动应用。
2. 拖拽文件到列表。
3. （可选）双击列表项设置页码范围 / 工作表选择。
4. 配置：
   - 水印 / 页码
   - 拆分/合并选项
   - 命名规则和输出文件夹
5. 点击 **开始转换 PDF**。

---

## 📦 安装

1. 从 **Releases** 下载最新版本。
2. 解压 `Office2PDF.zip`。
3. 运行 `Office2PDF.exe`（无需安装）。

---

## 🛠️ 从源码运行 / 构建可执行文件

### 1) 环境
- Python **3.9+**（Windows）
- 已安装 Microsoft Office 或支持 COM 的 WPS Office

### 2) 安装依赖
```bash
pip install pywin32 pypdf Pillow reportlab tkinterdnd2
```

### 3) 运行
```bash
python office2pdf_v5.py
```

### 4) 构建 `.exe`（PyInstaller）
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --name Office2PDF office2pdf_v5.py
```

构建完成后，可执行文件位于：
- `dist/Office2PDF.exe`

> 提示：某些环境下转换 PowerPoint 时需保持其可见。
> 如果 PowerPoint 转换失败，请避免强制 `Visible=False`。

---

## ⚖️ 许可证

MIT License — 详见 ./LICENSE
