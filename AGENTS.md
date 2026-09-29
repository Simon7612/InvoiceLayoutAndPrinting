# AGENTS.md

Windows 桌面工具：将多张发票/车票 PDF（也支持 OFD/XML 转 PDF）按多种 N-up 布局合成、预览并打印。技术栈：Python 3.12+ / PyQt6 / pypdf，Nuitka 打包单文件 exe。

## 模块与分层（扁平结构，无包目录）

- `main.py` — 薄入口。argparse：`-i`/`-o`/`--no-print`/`--gui`；无 `--gui` 且无 `-i` 时也直接进 GUI；批处理委托 `pipeline.process`
- `pipeline.py` — CLI/GUI 共用编排（**不依赖 Qt**，可无头测试）：`ComposeJob` + `run_compose_job()`（合成 → 写出 → 逐份打印，逐份失败只记日志）、`load_documents()`（读取/转换 + 车票识别聚合）、输出命名 `output_name()`/`LAYOUT_OUTPUT_NAMES`、CLI 批处理 `process()`（逐文件 2-up，命名 `<名>_2up.pdf`）
- `gui.py` — 主窗口组装（`MainWindow`、`run_gui()`、文件日志）；预览用 `QPdfDocument` + `QPdfView`（MultiPage + FitToWidth）
- `gui_workers.py` — 后台线程薄封装：`LoadWorker`（读取/转换+车票识别）、`ComposeWorker`（合成/写出/打印），只做信号转发与取消检查，编排都在 pipeline
- `gui_preferences.py` — QSettings 偏好读写（`Preferences.restore()/save()`，org/app 常量在此）
- `gui_widgets.py` — 拖拽导入区 `ImportDropArea`
- `layoutInvoice.py` — 纯排版/合成逻辑，核心 API：`compose_pages(pages, layout_mode, orientation, add_cutlines, grid=None)`（`orientation=None` 表示不做方向预处理，CLI 旧行为），模式：`one_up` / `two_up_vertical` / `two_up_horizontal` / `four_up`（=2×2） / `custom_grid`（需 `grid=[rows, cols]`）；通用实现 `grid_pages(pages, rows, cols)`（单元格取每组首页尺寸、逐页缩放填满）
- `readInvoice.py` — 文件收集与读取；`read_document()` 统一处理 PDF/OFD/XML，`detect_ticket_document()` 识别车票（中文关键词 1 命中即中、英文需 ≥2，防误报）
- `printInvoice.py` — 仅 Windows 的打印链（可注入策略列表）：`print_pdf(path, printer=None, strategies=None)`，缺省链路为指定打印机时 SumatraPDF `-print-to` 直印，否则 SumatraPDF `-print-dialog` → ShellExecuteW `print` 动词 → PowerShell → `os.startfile` → Edge/查看器兜底（`_default_strategies` 按优先级构建）；外部程序一律 `Popen` 非阻塞启动（`run` 会卡在打印窗口上）；`list_system_printers()` 走注册表

依赖方向：`main.py`/`gui.py` → pipeline → readInvoice + layoutInvoice + printInvoice；`pipeline.py` 不 import Qt；`layoutInvoice.py` 不 import GUI/IO，保持纯函数。

## 常用命令

```bash
uv sync                              # 安装依赖（uv 为主管理器）
uv run python main.py --gui          # 运行 GUI
uv run python main.py -i <pdf或目录> -o <输出目录> --no-print   # CLI 批处理（2-up 竖向，PDF/OFD/XML 均可）
make install | run | package | clean # Makefile；package 产出 dist/发票排版与打印.exe
```

- 测试：`uv run pytest`（`tests/`，夹具全部由 pypdf 程序化生成，不依赖真实发票；打印链用注入的假策略测试，不启动外部程序；GUI 路径用 `QT_QPA_PLATFORM=offscreen` 手工验证）。Linter：`uv run ruff check .`（务实最小集，见 pyproject）。CI 冒烟测试为 `python -c "import gui, pipeline, layoutInvoice, printInvoice, readInvoice"`。
- `Makefile` 的 `SHELL := cmd`（Windows cmd 而非 bash），内部通过 PowerShell 调用 `.venv\Scripts\python.exe`。在 Git Bash 里直接复制 make 内部命令可能不成立。

## 约定与坑

- **Windows 优先**：打印仅支持 Windows（非 Windows 会 raise）。CI（`.github/workflows/release.yml`）为 test/release 双 job，windows-latest 上用 Nuitka 构建。**发版由 pyproject 驱动**：把 version 变更合入 `main`，CI 读取版本并检查同名 Release——不存在则自动创建 `v{version}` 标签并发布，已存在则只跑测试；不要手动打 tag。发布说明由 `cliff.toml`（git-cliff）从 Conventional Commits 自动生成中文分组（feat/fix/perf/refactor，其余类型跳过），所以提交信息要规范；release 资产为 ASCII 名 `InvoiceLayoutAndPrinting-<版本>.exe`（GitHub 会把中文资产名规范化成 default.exe）。版本号唯一来源是 `pyproject.toml` 的 `version`（CI 用 tomllib 读取，Makefile 用 `findstr` 派生，改版本只改 pyproject 一行）。
- **文件名为 camelCase**（`layoutInvoice.py` 等），注释/README/UI 文案均为中文。
- **OFD/XML 依赖是可选导入**：`ofdparser`（首选，**依赖 numpy**——曾因未声明 numpy 而静默失效）/`easyofd`（备选，另需 cv2 处理图片）用 try/except 导入并有 `_HAS_OFDPARSER`/`_HAS_EASYOFD` 开关；`main.py` 和 `readInvoice.py` 顶部对这两个库的 `SyntaxWarning` 做了过滤（故 main.py 有 E402 豁免），改动导入顺序时别弄丢。
- **改合成逻辑时必须保持两件事**：1) 用每页 `cropbox` 对齐坐标系（不同来源 PDF 尺寸/裁剪框不同）；2) 电子印章注释定位——`merge_transformed_page` 会**原样**复制 `/Annots`（不做坐标变换），实际定位靠合并后对 `/Rect` 按"与内容完全相同的变换"修正（纯平移用 `_adjust_merged_annots`，含缩放用 `_transform_annots`），两处变换必须一致。回归测试在 `tests/test_layoutInvoice.py`（含 cropbox 原点偏移与宫格象限用例），回归会直接毁掉核心功能。
- **GUI 偏好用 QSettings 持久化**（org "Simon Chan" / app "InvoiceLayoutAndPrinting"，见 `gui_preferences.py`）：`closeEvent` 里 `Preferences.save()`、`__init__` 尾部 `Preferences.restore()`；程序化 `setChecked` 不触发 `idClicked`，`restore()` 末尾需手动调 `MainWindow.update_two_tile_caption()`/`update_grid_visibility()` 同步派生状态——新增随偏好变化的派生 UI 时，记得在 restore() 里补同步。
- **打包 exe 无控制台**：GUI 日志经 `run_gui()` 的 `RotatingFileHandler` 写入 `%LOCALAPPDATA%\InvoiceLayoutAndPrinting\logs\gui.log`（pipeline 的打印失败警告也随之落盘）。
- 输出命名固定为 `merged_1up.pdf` / `merged_2up_v.pdf` / `merged_2up_h.pdf` / `merged_4up.pdf` / `merged_grid{r}x{c}.pdf`（`pipeline.LAYOUT_OUTPUT_NAMES` + `pipeline.output_name()`，排版与打印按钮共用）；CLI 批处理为逐文件 `<名>_2up.pdf`（`pipeline.process`）。
- `.gitignore` 排除 `*.pdf` 和 `test/`：不要提交样例发票 PDF（tests 夹具有豁免 `!tests/**/*.pdf`）。
- Git：日常在 `dev` 分支开发，PR 合入 `main`。

## 改动前必读

- `README.md` 的「技术细节」与「常见问题」章节——记录了 cropbox/注释处理动机和拖拽排序黑色条带等已知问题。
