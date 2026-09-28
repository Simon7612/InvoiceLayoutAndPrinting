import ctypes
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from typing import List

from PyQt6.QtCore import Qt, QSize, QEvent
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QAbstractItemView,
    QLabel,
    QFileDialog,
    QGroupBox,
    QFormLayout,
    QGridLayout,
    QButtonGroup,
    QSpinBox,
    QCheckBox,
    QLineEdit,
    QComboBox,
    QMessageBox,
    QSplitter,
    QToolButton,
    QStyle,
)
from PyQt6.QtGui import QIcon
from readInvoice import collect_pdfs
from layoutInvoice import LayoutMode, Orientation
from printInvoice import print_pdf, list_system_printers
from pipeline import ComposeJob, output_name
from gui_widgets import ImportDropArea
from gui_workers import ComposeWorker, LoadWorker
from gui_preferences import Preferences
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("发票排版与打印")
        self.statusBar()
        self.setObjectName("MainWindow")
        splitter = QSplitter()
        left = QGroupBox("发票列表")
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.setSpacing(12)
        self.label_count = QLabel("已上传 0 个文件")
        self.btn_import = QPushButton("+ 添加发票")
        self.hint_left = QLabel("支持 PDF/OFD/XML \n拖拽发票到此处 或 点击上方按钮导入")
        self.hint_left.setAlignment(Qt.AlignmentFlag.AlignCenter)
        import_card = ImportDropArea(self.on_drop_files)
        import_card.setObjectName("ImportCard")
        card_layout = QVBoxLayout()
        card_layout.setSpacing(8)
        card_layout.addWidget(self.btn_import, alignment=Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.hint_left)
        import_card.setLayout(card_layout)
        self.list_files = QListWidget()
        try:
            self.list_files.setDragDropMode(QListWidget.DragDropMode.InternalMove)
            self.list_files.setDefaultDropAction(Qt.DropAction.MoveAction)
            self.list_files.setDragDropOverwriteMode(False)
            self.list_files.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            self.list_files.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        except Exception:
            pass
        left_layout.addWidget(self.label_count)
        left_layout.addWidget(import_card)
        left_layout.addWidget(self.list_files)
        left.setLayout(left_layout)
        center_box = QGroupBox("预览窗口")
        center_layout = QVBoxLayout()
        center_layout.setContentsMargins(8, 8, 8, 8)
        self.pdf_doc = QPdfDocument(self)
        self.pdf_view = QPdfView(self)
        try:
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        except Exception:
            pass
        try:
            self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        except Exception:
            pass
        # 预览工具条：缩放与页码
        bar = QHBoxLayout()
        self.btn_zoom_out = QToolButton()
        self.btn_zoom_out.setText("－")
        self.btn_zoom_out.setToolTip("缩小")
        self.btn_zoom_fit = QToolButton()
        self.btn_zoom_fit.setText("适应宽度")
        self.btn_zoom_in = QToolButton()
        self.btn_zoom_in.setText("＋")
        self.btn_zoom_in.setToolTip("放大")
        self.label_page = QLabel("未加载文档")
        bar.addWidget(self.btn_zoom_out)
        bar.addWidget(self.btn_zoom_fit)
        bar.addWidget(self.btn_zoom_in)
        bar.addStretch(1)
        bar.addWidget(self.label_page)
        center_layout.addLayout(bar)
        center_layout.addWidget(self.pdf_view)
        self.btn_zoom_in.clicked.connect(lambda: self._zoom_preview(1.25))
        self.btn_zoom_out.clicked.connect(lambda: self._zoom_preview(0.8))
        self.btn_zoom_fit.clicked.connect(self._zoom_fit_preview)
        try:
            self.pdf_view.pageNavigator().currentPageChanged.connect(
                lambda _page: self._update_page_label()
            )
        except Exception:
            pass
        center_box.setLayout(center_layout)
        right = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)
        layout_box = QGroupBox("排版方式")
        rb_layout = QVBoxLayout()
        # 1) 卡片式布局选择
        grid = QGridLayout()
        grid.setSpacing(8)

        def make_tile(text: str, enabled: bool=True) -> QToolButton:
            b = QToolButton()
            b.setCheckable(True)
            b.setEnabled(enabled)
            b.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            b.setText(text)
            b.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
            b.setIconSize(QSize(24, 24))
            b.setAutoRaise(False)
            b.setMinimumSize(QSize(120, 64))
            return b

        self.btn_layout_custom = make_tile("自定义\n自由定义布局", enabled=True)
        self.btn_layout_one = make_tile("单页\n一页一张")
        self.btn_layout_two_v = make_tile("双页\n上下布局")
        self.btn_layout_four = make_tile("四页\n2×2布局")

        grid.addWidget(self.btn_layout_custom, 0, 0)
        grid.addWidget(self.btn_layout_one,    0, 1)
        grid.addWidget(self.btn_layout_two_v,  1, 0)
        grid.addWidget(self.btn_layout_four,   1, 1)

        self.group_layout_tiles = QButtonGroup(self)
        for i, b in enumerate([self.btn_layout_custom, self.btn_layout_one, self.btn_layout_two_v, self.btn_layout_four]):
            self.group_layout_tiles.addButton(b, i)
        self.group_layout_tiles.setExclusive(True)
        self.btn_layout_one.setChecked(True)  # 默认单页一张

        rb_layout.addLayout(grid)

        # 2) 纸张方向（纵向/横向）
        h_orient = QHBoxLayout()
        self.btn_portrait = QToolButton()
        self.btn_portrait.setCheckable(True)
        self.btn_portrait.setText("纵向")
        self.btn_portrait.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowUp))
        self.btn_portrait.setIconSize(QSize(20, 20))
        self.btn_landscape = QToolButton()
        self.btn_landscape.setCheckable(True)
        self.btn_landscape.setText("横向")
        self.btn_landscape.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight))
        self.btn_landscape.setIconSize(QSize(20, 20))
        self.group_orient = QButtonGroup(self)
        self.group_orient.addButton(self.btn_portrait, 0)
        self.group_orient.addButton(self.btn_landscape, 1)
        self.group_orient.setExclusive(True)
        self.btn_portrait.setChecked(True)
        h_orient.addWidget(QLabel("纸张方向"))
        h_orient.addStretch(1)
        h_orient.addWidget(self.btn_portrait)
        h_orient.addWidget(self.btn_landscape)
        rb_layout.addLayout(h_orient)

        # 根据方向动态更新“双页”卡片的副标题：纵向→上下布局；横向→左右布局
        self.update_two_tile_caption()
        # 方向变化时刷新
        try:
            self.group_orient.idClicked.connect(lambda _id: self.update_two_tile_caption())
        except Exception:
            # 兜底：直接监听两个按钮的toggled
            try:
                self.btn_portrait.toggled.connect(lambda _checked: self.update_two_tile_caption())
                self.btn_landscape.toggled.connect(lambda _checked: self.update_two_tile_caption())
            except Exception:
                pass

        # 3) 自定义网格：行 × 列（每页张数 = 行×列，逐页缩放填入单元格）
        h_grid = QHBoxLayout()
        h_grid.addWidget(QLabel("网格 行 × 列"))
        self.spin_grid_rows = QSpinBox()
        self.spin_grid_rows.setRange(1, 5)
        self.spin_grid_rows.setValue(2)
        self.spin_grid_cols = QSpinBox()
        self.spin_grid_cols.setRange(1, 5)
        self.spin_grid_cols.setValue(2)
        for s in (self.spin_grid_rows, self.spin_grid_cols):
            s.setMinimumWidth(60)
        h_grid.addStretch(1)
        h_grid.addWidget(self.spin_grid_rows)
        h_grid.addWidget(QLabel("×"))
        h_grid.addWidget(self.spin_grid_cols)
        # 包一层，便于动态隐藏/显示
        self.grid_wrap = QWidget()
        self.grid_wrap.setLayout(h_grid)
        rb_layout.addWidget(self.grid_wrap)

        # 4) 切割线
        self.chk_cutline = QCheckBox("显示切割线")
        rb_layout.addWidget(self.chk_cutline)
        layout_box.setLayout(rb_layout)
        opt_box = QGroupBox("选项")
        form = QFormLayout()
        self.spin_copies = QSpinBox()
        self.spin_copies.setMinimum(1)
        self.spin_copies.setValue(1)
        self.chk_print = QCheckBox("排版后打印")
        self.line_out = QLineEdit()
        self.line_out.setPlaceholderText("输出目录，留空使用源目录")
        self.btn_out = QPushButton("📁 选择输出目录")
        form.addRow("份数", self.spin_copies)
        form.addRow("打印", self.chk_print)
        h_out = QHBoxLayout()
        h_out.addWidget(self.line_out)
        h_out.addWidget(self.btn_out)
        w_out = QWidget()
        w_out.setLayout(h_out)
        form.addRow("输出目录", w_out)
        opt_box.setLayout(form)
        self.combo_printer = QComboBox()
        self.combo_printer.addItem("默认打印机")
        try:
            for name in list_system_printers():
                self.combo_printer.addItem(name)
        except Exception:
            pass
        # 车票选项
        ticket_box = QGroupBox("车票选项")
        ticket_form = QFormLayout()
        self.chk_ticket_duplicate = QCheckBox("一页重复两张")
        ticket_form.addRow("重复排版", self.chk_ticket_duplicate)
        ticket_box.setLayout(ticket_form)
        btns = QHBoxLayout()
        btns.setSpacing(12)
        self.btn_layout = QPushButton("🧩 排版")
        self.btn_print = QPushButton("🖨 打印")
        btns.addWidget(self.btn_layout)
        btns.addWidget(self.btn_print)
        right_wrap = QGroupBox("打印设置")
        right_inner = QVBoxLayout()
        right_inner.addWidget(layout_box)
        right_inner.addWidget(opt_box)
        right_inner.addWidget(self.combo_printer)
        right_inner.addWidget(ticket_box)
        r_btns = QWidget()
        r_btns.setLayout(btns)
        right_inner.addWidget(r_btns)
        right_wrap.setLayout(right_inner)
        right_layout.addWidget(right_wrap)
        right_layout.addStretch(1)
        right.setLayout(right_layout)
        splitter.addWidget(left)
        splitter.addWidget(center_box)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 1)
        self.setCentralWidget(splitter)
        try:
            self.list_files.installEventFilter(self)
        except Exception:
            pass
        self.btn_import.clicked.connect(self.on_import)
        self.btn_out.clicked.connect(self.on_choose_out)
        self.btn_layout.clicked.connect(self.on_layout)
        self.btn_print.clicked.connect(self.on_print)
        # 动态显示/隐藏“网格 行×列”：仅自定义卡片时显示
        self.update_grid_visibility()
        self.group_layout_tiles.idClicked.connect(lambda _id: self.update_grid_visibility())
        # 偏好恢复（窗口尺寸、布局、打印设置等）；内部会同步派生 UI 状态
        self._prefs = Preferences(self)
        self._prefs.restore()

    def eventFilter(self, obj, event):
        try:
            if obj is self.list_files and event.type() == QEvent.Type.Resize:
                self._update_list_item_widths()
        except Exception:
            pass
        return False

    def update_two_tile_caption(self) -> None:
        """按纸张方向同步“双页”卡片副标题（纵向→上下布局；横向→左右布局）。"""
        if self.btn_landscape.isChecked():
            self.btn_layout_two_v.setText("双页\n左右布局")
        else:
            self.btn_layout_two_v.setText("双页\n上下布局")

    def update_grid_visibility(self) -> None:
        """仅自定义布局卡片（id=0）时显示“网格 行×列”。"""
        self.grid_wrap.setVisible(self.group_layout_tiles.checkedId() == 0)

    def _update_list_item_widths(self) -> None:
        vw = self.list_files.viewport().width()
        for i in range(self.list_files.count()):
            it = self.list_files.item(i)
            w = self.list_files.itemWidget(it)
            if not w:
                continue
            try:
                lbl = w.findChild(QLabel)
                if lbl:
                    fm = lbl.fontMetrics()
                    maxw = max(40, vw - 64)
                    full_path = it.data(Qt.ItemDataRole.UserRole)
                    name = os.path.basename(full_path) if isinstance(full_path, str) else lbl.text()
                    lbl.setText(fm.elidedText(name, Qt.TextElideMode.ElideMiddle, maxw))
            except Exception:
                pass
    def on_drop_files(self, paths: List[str]):
        self.add_paths(paths)
    def add_paths(self, paths: List[str]):
        files: List[str] = []
        for p in paths:
            files.extend(collect_pdfs(p))
        existing = set(self.get_files())
        for f in files:
            if f not in existing:
                self.add_list_item(f)
        self.label_count.setText(f"已上传 {self.list_files.count()} 个文件")
    def add_list_item(self, full_path: str):
        name = os.path.basename(full_path)
        it = QListWidgetItem()
        it.setData(Qt.ItemDataRole.UserRole, full_path)
        try:
            it.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsDragEnabled)
        except Exception:
            pass
        w = QWidget()
        w.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        w.setStyleSheet("background:#ffffff;")
        hl = QHBoxLayout()
        hl.setContentsMargins(8, 4, 8, 4)
        lbl = QLabel(name)
        lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        fm = lbl.fontMetrics()
        vw = self.list_files.viewport().width()
        maxw = max(40, vw - 64)
        lbl.setText(fm.elidedText(name, Qt.TextElideMode.ElideMiddle, maxw))
        lbl.setToolTip(full_path)
        btn = QToolButton()
        btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCloseButton))
        btn.setAutoRaise(True)
        btn.setIconSize(QSize(16, 16))
        btn.setToolTip("移除")
        btn.clicked.connect(lambda: self.remove_list_item(it))
        hl.addWidget(lbl)
        hl.addStretch(1)
        btn_wrap = QWidget()
        btn_wrap.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        btn_wrap.setStyleSheet("background:#ffffff;")
        btn_wrap.setFixedWidth(28)
        bl = QHBoxLayout()
        bl.setContentsMargins(0, 0, 0, 0)
        bl.addStretch(1)
        bl.addWidget(btn)
        bl.addStretch(1)
        btn_wrap.setLayout(bl)
        hl.addWidget(btn_wrap)
        w.setLayout(hl)
        it.setSizeHint(w.sizeHint())
        self.list_files.addItem(it)
        self.list_files.setItemWidget(it, w)
    def remove_list_item(self, item: QListWidgetItem):
        row = self.list_files.row(item)
        if row >= 0:
            self.list_files.takeItem(row)
            self.label_count.setText(f"已上传 {self.list_files.count()} 个文件")
    def get_files(self) -> List[str]:
        out: List[str] = []
        for i in range(self.list_files.count()):
            it = self.list_files.item(i)
            p = it.data(Qt.ItemDataRole.UserRole)
            out.append(p if isinstance(p, str) else it.text())
        return out
    def on_import(self):
        dlg = QFileDialog(self)
        dlg.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dlg.setNameFilter("文档 (*.pdf *.ofd *.xml)")
        if dlg.exec():
            self.add_paths(dlg.selectedFiles())
    def on_choose_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.line_out.setText(d)
    def on_layout(self):
        files = self.get_files()
        if not files:
            QMessageBox.warning(self, "提示", "请先导入发票")
            return
        # 主线程先取齐所有设置，重活交给后台线程
        self._layout_files = files
        self._layout_out_dir = self.line_out.text().strip() or None
        self._layout_do_print = self.chk_print.isChecked()
        self._layout_copies = self.spin_copies.value()
        self._layout_printer = self._selected_printer()
        self.set_busy(True)
        self.statusBar().showMessage("正在读取/转换文件…")
        # 阶段一：读取/转换 + 车票识别（OFD/XML 转换耗时）
        self._load_worker = LoadWorker(files, parent=self)
        self._load_worker.progress.connect(self._on_load_progress)
        self._load_worker.ticket_detected.connect(self._on_ticket_detected)
        self._load_worker.loaded.connect(self._on_files_loaded)
        self._load_worker.failed.connect(self._on_layout_failed)
        self._load_worker.start()

    def _resolve_layout_mode(self) -> str:
        """根据右侧卡片与方向选择解析布局模式。

        - 单页卡片：固定 ONE_UP
        - 双页卡片：纵向→上下，横向→左右
        - 四页卡片：固定 FOUR_UP
        - 自定义卡片：按“网格 行×列”决定
        """
        checked_tile = self.group_layout_tiles.checkedId()
        if checked_tile == 1:
            return LayoutMode.ONE_UP
        if checked_tile == 2:
            return LayoutMode.TWO_UP_VERTICAL if self.btn_portrait.isChecked() else LayoutMode.TWO_UP_HORIZONTAL
        if checked_tile == 3:
            return LayoutMode.FOUR_UP
        return LayoutMode.CUSTOM_GRID

    def _output_name(self, mode: str) -> str:
        return output_name(mode, [self.spin_grid_rows.value(), self.spin_grid_cols.value()])

    def _selected_printer(self) -> str | None:
        """下拉框第一项为“默认打印机”（弹窗），其余为直印目标。"""
        if self.combo_printer.currentIndex() <= 0:
            return None
        return self.combo_printer.currentText()

    def _zoom_preview(self, factor: float):
        try:
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
            self.pdf_view.setZoomFactor(self.pdf_view.zoomFactor() * factor)
        except Exception:
            pass

    def _zoom_fit_preview(self):
        try:
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        except Exception:
            pass

    def _update_page_label(self):
        try:
            n = self.pdf_doc.pageCount()
            if n <= 0:
                self.label_page.setText("未加载文档")
                return
            current = self.pdf_view.pageNavigator().currentPage()
            self.label_page.setText(f"第 {current + 1} / {n} 页")
        except Exception:
            pass

    def _on_load_progress(self, i: int, total: int, name: str):
        self.statusBar().showMessage(f"正在读取/转换文件（{i}/{total}）：{name}")

    def _on_ticket_detected(self, any_ticket: bool, suggested: str):
        if not any_ticket:
            return
        reply = QMessageBox.question(
            self, "检测到车票",
            "识别到车票样式。是否启用“一页重复两张”并预设为双页排版？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self.statusBar().showMessage("检测到车票：保持当前排版设置", 5000)
            return
        # 自动勾选“重复两张”并预设 2-up 与建议方向
        self.chk_ticket_duplicate.setChecked(True)
        self.btn_layout_two_v.setChecked(True)
        self.btn_layout_one.setChecked(False)
        self.btn_layout_four.setChecked(False)
        if suggested == "landscape":
            self.btn_landscape.setChecked(True)
            self.btn_portrait.setChecked(False)
        elif suggested == "portrait":
            self.btn_portrait.setChecked(True)
            self.btn_landscape.setChecked(False)
        self.update_two_tile_caption()
        # 隐藏“网格 行×列”保持与非自定义一致
        self.grid_wrap.setVisible(False)
        self.statusBar().showMessage("检测到车票：已启用重复两张并预设为 2-up", 5000)

    def _on_files_loaded(self, readers: list):
        pages: List = []
        for _src, r in readers:
            pages.extend(list(r.pages))
        if self.chk_ticket_duplicate.isChecked():
            # 将每页复制一份再进行左右或上下 2-up
            dup_pages: List = []
            for p in pages:
                dup_pages.extend([p, p])
            pages = dup_pages
        # 布局与方向映射见 _resolve_layout_mode
        mode = self._resolve_layout_mode()
        grid = None
        if mode == LayoutMode.CUSTOM_GRID:
            grid = [self.spin_grid_rows.value(), self.spin_grid_cols.value()]
        od = self._layout_out_dir or os.path.dirname(self._layout_files[0])
        job = ComposeJob(
            pages=pages,
            mode=mode,
            orientation=Orientation.PORTRAIT if self.btn_portrait.isChecked() else Orientation.LANDSCAPE,
            add_cutlines=self.chk_cutline.isChecked(),
            grid=grid,
            out_path=os.path.join(od, self._output_name(mode)),
            do_print=self._layout_do_print,
            copies=self._layout_copies,
            printer=self._layout_printer,
        )
        # 阶段二：合成、写出与打印
        self._compose_worker = ComposeWorker(job, parent=self)
        self._compose_worker.progress.connect(self.statusBar().showMessage)
        self._compose_worker.composed.connect(self._on_composed)
        self._compose_worker.failed.connect(self._on_layout_failed)
        self._compose_worker.start()

    def _on_composed(self, out_path: str):
        self.set_busy(False)
        self.statusBar().clearMessage()
        try:
            self.load_preview(out_path)
        except Exception:
            pass
        QMessageBox.information(self, "完成", f"已生成 {out_path}")

    def _on_layout_failed(self, message: str):
        self.set_busy(False)
        self.statusBar().clearMessage()
        QMessageBox.critical(self, "排版失败", message)
    def on_print(self):
        files = self.get_files()
        if not files:
            QMessageBox.warning(self, "提示", "请先导入发票")
            return
        out_dir = self.line_out.text().strip() or None
        od = out_dir or os.path.dirname(files[0])
        # 按当前选择的布局模式定位排版输出文件
        target = os.path.join(od, self._output_name(self._resolve_layout_mode()))
        if not os.path.exists(target):
            QMessageBox.information(self, "提示", "未找到排版后的文件，请先排版")
            return
        copies = self.spin_copies.value()
        printer = self._selected_printer()
        self.set_busy(True)
        self.statusBar().showMessage("正在打开打印对话框…")
        try:
            for _ in range(copies):
                try:
                    print_pdf(target, printer)
                except Exception as e:
                    logging.getLogger(__name__).warning("打印失败 %s: %s", target, e)
        finally:
            self.set_busy(False)
            self.statusBar().clearMessage()

    def set_busy(self, busy: bool):
        for b in [self.btn_layout, self.btn_print, self.btn_import, self.btn_out]:
            b.setEnabled(not busy)
        if busy:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        else:
            QApplication.restoreOverrideCursor()

    def closeEvent(self, event):
        self._prefs.save()
        # 后台线程仍在读取/排版时，先请求中断并等待当前文件处理完，
        # 避免线程随窗口销毁导致的崩溃
        for attr in ("_load_worker", "_compose_worker"):
            worker = getattr(self, attr, None)
            if worker is not None and worker.isRunning():
                worker.requestInterruption()
                worker.wait()
        event.accept()

    def load_preview(self, path: str):
        self.pdf_doc.load(path)
        self.pdf_view.setDocument(self.pdf_doc)
        try:
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        except Exception:
            pass
        try:
            self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        except Exception:
            pass
        self._update_page_label()

def _setup_file_logging():
    """打包后的 exe 没有控制台，日志写入本地文件便于排查印章/打印问题。"""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    log_dir = os.path.join(base, "InvoiceLayoutAndPrinting", "logs")
    try:
        os.makedirs(log_dir, exist_ok=True)
        handler = RotatingFileHandler(
            os.path.join(log_dir, "gui.log"),
            maxBytes=1_000_000, backupCount=3, encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        root = logging.getLogger()
        root.setLevel(logging.INFO)
        root.addHandler(handler)
    except Exception:
        pass  # 日志不可用时不能阻断 GUI 启动

def run_gui():
    _setup_file_logging()
    logging.getLogger(__name__).info("启动 GUI")
    app = QApplication([])
    # Apply app icon for taskbar/titlebar
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("InvoiceLayoutAndPrinting.SimonChan")
    except Exception:
        pass
    icon_path = None
    try:
        if sys.executable.lower().endswith(".exe"):
            p = os.path.join(os.path.dirname(sys.executable), "icon.ico")
            if os.path.exists(p):
                icon_path = p
    except Exception:
        pass
    if not icon_path:
        p2 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        if os.path.exists(p2):
            icon_path = p2
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
    w = MainWindow()
    if icon_path:
        w.setWindowIcon(QIcon(icon_path))
    w.resize(1200, 700)
    w.show()
    app.setStyle("Fusion")
    app.setStyleSheet(
        """
        #MainWindow { background: #ffffff; }
        QListWidget { border: 1px solid #e5e7eb; border-radius: 8px; padding: 6px; background: #ffffff; }
        QListWidget::item:selected { background: #e3f2fd; color: #111827; }
        QGroupBox { border: 1px solid #dbe1ea; border-radius: 10px; margin-top: 12px; }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #374151; }
        QPushButton { background: #3b82f6; color: #fff; border: none; padding: 8px 14px; border-radius: 8px; }
        QPushButton:hover { background: #2563eb; }
        QPushButton:disabled { background: #93c5fd; }
        QLineEdit { border: 1px solid #e5e7eb; border-radius: 6px; padding: 6px; }
        QSpinBox, QComboBox { border: 1px solid #e5e7eb; border-radius: 6px; padding: 4px; }
        #ImportCard { border: 1px dashed #93c5fd; border-radius: 12px; padding: 16px; }
        """
    )
    app.exec()
