# 发票排版与打印（InvoiceLayoutAndPrinting）

一个用于将多张 PDF 发票按“上下双页”方式自动合并、预览与打印的桌面工具。基于 Python、PyQt6 与 pypdf，支持一键打包为单文件可执行程序（Nuitka）。

## 功能概述
- 导入单个 PDF 或目录中的所有 PDF 文件（拖拽或文件选择）
- 文件列表仅显示文件名，支持拖拽排序，右侧系统风格“关闭”图标一键移除
- 排版方式：双页竖向合并（两页拼成一页，上下排列）
- 预览窗口支持多页滚动查看，自动适配宽度
- 打印：支持系统打印对话框，或调用 Edge 打印（如可用）
- 兼容不同尺寸与裁剪框的发票 PDF，修复印章（PDF 注释）在合成后位置偏移的问题

## 快速开始

### 环境准备
项目使用 `uv` 管理依赖与运行；若未安装 `uv`，可使用 Python 原生虚拟环境与 pip。

- 使用 `uv`：
  - 安装依赖：`uv sync`
  - 运行 GUI：`uv run python main.py --gui`
  - 命令行排版：`uv run python main.py -i <PDF或目录> -o <输出目录> --no-print`

- 使用虚拟环境与 pip：
  - 创建虚拟环境：`python -m venv .venv`
  - 安装依赖：`.venv\Scripts\python -m pip install -U pip nuitka pyqt6 pypdf`
  - 运行 GUI：`.venv\Scripts\python main.py --gui`

### 项目脚本
- `Makefile`（Windows 环境适配，内部调用 PowerShell 执行 Python）：
  - 安装依赖：`make install`
  - 打包可执行文件：`make package`（生成 `dist/发票排版与打印.exe`）
  - 运行程序：`make run`
  - 清理构建产物：`make clean`

## 使用说明
- 添加发票：拖拽 PDF 到左侧卡片或点击“+ 添加发票”选择文件/目录
- 文件列表：
  - 仅显示文件名（悬停显示完整路径）
  - 右侧“关闭”图标可移除条目
  - 支持拖拽排序，列表当前顺序决定合并后的页序
- 排版：点击“🧩 排版”生成合并后的 PDF（默认输出到源目录，或指定输出目录）
- 打印：勾选“排版后打印”，或在右侧点击“🖨 打印”
- 预览：排版完成后自动加载合并文件，多页滚动查看

## 技术细节
- 合成逻辑在 `layoutInvoice.py`：
  - `two_up_vertical_pages(pages)` 按两页一组竖向合成；宽度取两页最大值，高度为两页高度和
  - 使用每页的 `cropbox` 对齐坐标系，保证不同来源 PDF 的布局一致
  - 对 PDF 注释（`/Annots`，如电子印章）进行同步平移与复制，确保印章位置在合成后仍处于票头处
- GUI 在 `gui.py`：
  - 文件列表使用 `QListWidget` 自定义行控件，支持拖拽排序、系统图标删除按钮
  - 预览使用 `QPdfDocument` + `QPdfView`，启用 `MultiPage` 模式与 `FitToWidth`
- 打印在 `printInvoice.py`：
  - 优先尝试 Edge 的打印对话框；不可用则调用 Windows Shell 打印或打开默认查看器

## 常见问题
- 路径包含特殊字符（如 `&`）：命令行中会被当作分隔符；本项目的 `Makefile` 已通过在 PowerShell 中调用虚拟环境 Python 并对参数加引号进行规避
- 预览只显示第一页：已启用 `QPdfView.PageMode.MultiPage`，滚动可见所有页
- 拖拽排序时底部出现黑色条带：
  - 原因：窗口初始化后列表行控件宽度未与视口完全同步，拖拽重绘时暴露未绘制区域
  - 处理：关闭横向滚动、逐像素滚动、为行容器设置背景，并在列表 `Resize` 时重新计算标签省略宽度；最大化或适当调整窗口尺寸也可使现象消失

## 目录结构
```
InvoiceLayoutAndPrinting/
├─ gui.py                 # 图形界面
├─ main.py                # CLI 与 GUI 入口
├─ layoutInvoice.py       # 合成与注释处理逻辑
├─ printInvoice.py        # 打印实现
├─ readInvoice.py         # 读取与收集 PDF
├─ Makefile               # 构建与打包
├─ pyproject.toml         # 依赖与项目配置
├─ uv.lock                # uv 锁文件
├─ icon.ico               # 应用图标
└─ README.md              # 项目说明
```

## 打包说明（Nuitka）
- 单文件打包、禁用控制台、启用 PyQt6 插件，保留图标与元信息；构建后产物为 `dist/发票排版与打印.exe`
- 执行：`make package`

## 许可
未设置专用许可证；如需开源许可证可补充（例如 MIT）。

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Simon7612/InvoiceLayoutAndPrinting&type=date&legend=top-left)](https://www.star-history.com/#Simon7612/InvoiceLayoutAndPrinting&type=date&legend=top-left)

[v0.0.2 计划与实现]

- 页面排版扩展：
  - 纸张方向可选：`纵向`/`横向`
  - 布局模式：`一页一张`、`一页两张（上下）`、`一页两张（左右）`、`一页四张（2×2，固定横向）`
  - 切割线：支持在合成页中间添加切割线（上下布局绘制水平线、左右布局绘制竖线、四宫格绘制十字线）
- 车票打印：在“车票选项”中可选“`一页重复两张`”，用于同一车票的两连版排版（复用 2-up 算法）
- OFD/XML 支持：内置纯 Python 转换（无需外部工具）
  - OFD：优先使用 `ofdparser` 将 OFD 转为 PDF；若安装了 `easyofd` 也可作为备选
    - 安装：`uv add ofdparser`（或 `uv add easyofd`）
  - XML：支持解析 XML 中内嵌的 `base64` 负载（PDF 或 OFD）并转换为 PDF
    - 若 XML 不包含内嵌 PDF/OFD 数据，则当前版本无法直接渲染
  - GUI 导入已支持 `*.pdf *.ofd *.xml`；在未安装上述库时，OFD/XML 会提示安装方式

### 使用说明（GUI）
- 左侧导入列表添加发票/票据（可多选/拖拽），支持 `PDF/OFD/XML`
- 右侧选择：`布局模式`、`纸张方向`、`添加中间切割线`、以及可选的`车票重复两张`
- 点击“排版”生成输出，并在中间预览窗口展示；可选“排版后打印”直接调用系统打印
- 输出文件命名示例：`merged_1up.pdf`、`merged_2up_v.pdf`、`merged_2up_h.pdf`、`merged_4up.pdf`

### 命令行（CLI）
- 原 CLI 仍支持 2-up（上下）批处理；新版高级布局优先通过 GUI 使用
