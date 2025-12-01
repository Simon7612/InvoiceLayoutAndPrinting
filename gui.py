import os
from typing import List
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
    QRadioButton,
    QFormLayout,
    QSpinBox,
    QCheckBox,
    QLineEdit,
    QComboBox,
    QMessageBox,
    QSplitter,
    QToolButton,
    QStyle,
)
from PyQt6.QtCore import Qt, QSize, QEvent
from PyQt6.QtGui import QIcon
import ctypes, sys
from readInvoice import collect_pdfs, read_pdf, read_document, detect_ticket_document
from layoutInvoice import (
    two_up_vertical,
    two_up_vertical_pages,
    write_writer,
    compose_pages,
    LayoutMode,
    Orientation,
)
from printInvoice import print_pdf
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView

class DropArea(QWidget):
    def __init__(self, on_dropped):
        super().__init__()
        self.setAcceptDrops(True)
        self.label = QLabel("预览窗口\n请从左侧上传发票文件查看预览")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        self.setLayout(layout)
        self.on_dropped = on_dropped
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
    def dropEvent(self, e):
        paths: List[str] = []
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p:
                paths.append(p)
        if paths:
            self.on_dropped(paths)

class ImportDropArea(QWidget):
    def __init__(self, on_dropped):
        super().__init__()
        self.setAcceptDrops(True)
        self.on_dropped = on_dropped
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
    def dropEvent(self, e):
        paths: List[str] = []
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if p:
                paths.append(p)
        if paths:
            self.on_dropped(paths)

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
        center_layout.addWidget(self.pdf_view)
        center_box.setLayout(center_layout)
        right = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)
        layout_box = QGroupBox("排版方式")
        rb_layout = QVBoxLayout()
        # 1) 卡片式布局选择
        from PyQt6.QtWidgets import QGridLayout, QButtonGroup
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
        self.btn_portrait = QToolButton(); self.btn_portrait.setCheckable(True); self.btn_portrait.setText("纵向"); self.btn_portrait.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowUp)); self.btn_portrait.setIconSize(QSize(20,20))
        self.btn_landscape = QToolButton(); self.btn_landscape.setCheckable(True); self.btn_landscape.setText("横向"); self.btn_landscape.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight)); self.btn_landscape.setIconSize(QSize(20,20))
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
        def update_two_tile_caption():
            if self.btn_landscape.isChecked():
                self.btn_layout_two_v.setText("双页\n左右布局")
            else:
                self.btn_layout_two_v.setText("双页\n上下布局")
        # 初始刷新一次
        update_two_tile_caption()
        # 方向变化时刷新
        try:
            self.group_orient.idClicked.connect(lambda _id: update_two_tile_caption())
        except Exception:
            # 兜底：直接监听两个按钮的toggled
            try:
                self.btn_portrait.toggled.connect(lambda _checked: update_two_tile_caption())
                self.btn_landscape.toggled.connect(lambda _checked: update_two_tile_caption())
            except Exception:
                pass

        # 3) 每页发票数（影响布局：1张=ONE_UP；2张(左右)=TWO_UP_HORIZONTAL；4张=FOUR_UP）
        h_count = QHBoxLayout()
        h_count.addWidget(QLabel("每页发票数"))
        self.btn_count_1 = QToolButton(); self.btn_count_1.setCheckable(True); self.btn_count_1.setText("1张")
        self.btn_count_2h = QToolButton(); self.btn_count_2h.setCheckable(True); self.btn_count_2h.setText("2张")
        self.btn_count_4 = QToolButton(); self.btn_count_4.setCheckable(True); self.btn_count_4.setText("4张 (2×2)")
        for b in [self.btn_count_1, self.btn_count_2h, self.btn_count_4]:
            b.setMinimumWidth(90)
        self.group_count = QButtonGroup(self)
        self.group_count.addButton(self.btn_count_1, 1)
        self.group_count.addButton(self.btn_count_2h, 2)
        self.group_count.addButton(self.btn_count_4, 4)
        self.group_count.setExclusive(True)
        self.btn_count_2h.setChecked(False)
        self.btn_count_1.setChecked(False)
        self.btn_count_4.setChecked(False)
        h_count.addStretch(1)
        h_count.addWidget(self.btn_count_1)
        h_count.addWidget(self.btn_count_2h)
        h_count.addWidget(self.btn_count_4)
        # 包一层，便于动态隐藏/显示
        self.count_wrap = QWidget()
        self.count_wrap.setLayout(h_count)
        rb_layout.addWidget(self.count_wrap)

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
        # 同步：每页发票数 ↔ 卡片布局
        def sync_from_count(id_: int):
            # 任何“每页发票数”的选择都切换到“自定义”卡片，并显示该区域
            self.btn_layout_custom.setChecked(True)
            self.btn_layout_one.setChecked(False)
            self.btn_layout_two_v.setChecked(False)
            self.btn_layout_four.setChecked(False)
            update_count_visibility()
        self.group_count.idClicked.connect(sync_from_count)

        def sync_from_tiles(id_: int):
            # group_layout_tiles: 0=自定义(禁用),1=单页,2=双页上下,3=四页
            if id_ == 1:
                self.btn_count_1.setChecked(True)
            elif id_ == 2:
                self.btn_count_2h.setChecked(True)
            elif id_ == 3:
                self.btn_count_4.setChecked(True)
        self.group_layout_tiles.idClicked.connect(sync_from_tiles)

        # 动态显示/隐藏“每页发票数”：仅自定义时显示
        def update_count_visibility():
            checked_id = self.group_layout_tiles.checkedId()
            # 0=自定义(禁用按钮但可将来启用)，1=单页，2=双页上下，3=四页
            show = (checked_id == 0)
            self.count_wrap.setVisible(show)
        # 初始状态
        update_count_visibility()
        # 在卡片选择变化时更新
        self.group_layout_tiles.idClicked.connect(lambda _id: update_count_visibility())
        
    def eventFilter(self, obj, event):
        try:
            if obj is self.list_files and event.type() == QEvent.Type.Resize:
                self._update_list_item_widths()
        except Exception:
            pass
        return False

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
        out_dir = self.line_out.text().strip() or None
        do_print = self.chk_print.isChecked()
        copies = self.spin_copies.value()
        generated: List[str] = []
        self.set_busy(True)
        self.statusBar().showMessage("正在排版与输出…")
        try:
            pages: List = []
            any_ticket = False
            suggested_orient: str | None = None
            for src in files:
                try:
                    r = read_document(src)
                except Exception:
                    # 回退到 PDF
                    r = read_pdf(src)
                # 车票识别（按文件）
                try:
                    is_ticket, orient_hint = detect_ticket_document(r)
                    if is_ticket:
                        any_ticket = True
                        # 记录一个方向建议（若多文件不一致，保留第一个）
                        if not suggested_orient and orient_hint:
                            suggested_orient = orient_hint
                except Exception:
                    pass
                pages.extend(list(r.pages))
            # 车票重复排版（自动识别触发）
            if any_ticket:
                self.chk_ticket_duplicate.setChecked(True)
                # 预设“双页”卡片
                self.btn_layout_two_v.setChecked(True)
                self.btn_layout_one.setChecked(False)
                self.btn_layout_four.setChecked(False)
                # 根据建议方向设置
                if suggested_orient == "landscape":
                    self.btn_landscape.setChecked(True)
                    self.btn_portrait.setChecked(False)
                elif suggested_orient == "portrait":
                    self.btn_portrait.setChecked(True)
                    self.btn_landscape.setChecked(False)
                # 同步更新“双页”卡片文字
                try:
                    update_two_tile_caption()
                except Exception:
                    pass
                # 隐藏“每页发票数”保持与非自定义一致
                try:
                    self.count_wrap.setVisible(False)
                except Exception:
                    pass
                # 状态提示
                try:
                    self.statusBar().showMessage("检测到车票：已启用重复两张并预设为 2-up", 5000)
                except Exception:
                    pass
            ticket_duplicate = self.chk_ticket_duplicate.isChecked()
            if ticket_duplicate:
                # 将每页复制一份再进行左右或上下 2-up
                dup_pages = []
                for p in pages:
                    dup_pages.extend([p, p])
                pages = dup_pages
            # 布局与方向（基于按钮组）
            count_id = self.group_count.checkedId()
            # 响应式映射：
            # - 单页：仅看方向开关（仍用于旋转页以适应纸张），模式固定 ONE_UP
            # - 双页：根据方向决定上下/左右；纵向→上下，横向→左右
            # - 四页：固定 FOUR_UP
            # - 自定义：显示“每页发票数”，用其决定 1/2(左右)/4

            checked_tile = self.group_layout_tiles.checkedId()
            if checked_tile == 1:  # 单页
                mode = LayoutMode.ONE_UP
            elif checked_tile == 2:  # 双页
                mode = LayoutMode.TWO_UP_VERTICAL if self.btn_portrait.isChecked() else LayoutMode.TWO_UP_HORIZONTAL
            elif checked_tile == 3:  # 四页
                mode = LayoutMode.FOUR_UP
            else:  # 自定义
                if count_id == 1:
                    mode = LayoutMode.ONE_UP
                elif count_id == 2:
                    mode = LayoutMode.TWO_UP_HORIZONTAL
                elif count_id == 4:
                    mode = LayoutMode.FOUR_UP
                else:
                    mode = LayoutMode.ONE_UP
            orient = Orientation.PORTRAIT if self.btn_portrait.isChecked() else Orientation.LANDSCAPE
            add_cut = self.chk_cutline.isChecked()
            writer = compose_pages(pages, mode, orient, add_cut)
            # 输出文件名根据模式命名
            base_map = {
                LayoutMode.ONE_UP: "merged_1up.pdf",
                LayoutMode.TWO_UP_VERTICAL: "merged_2up_v.pdf",
                LayoutMode.TWO_UP_HORIZONTAL: "merged_2up_h.pdf",
                LayoutMode.FOUR_UP: "merged_4up.pdf",
            }
            base_name = base_map[mode]
            od = out_dir or os.path.dirname(files[0])
            out_path = os.path.join(od, base_name)
            write_writer(writer, out_path)
            generated.append(out_path)
            self.load_preview(out_path)
            if do_print:
                for _ in range(copies):
                    try:
                        print_pdf(out_path)
                    except Exception:
                        pass
        finally:
            self.set_busy(False)
            self.statusBar().clearMessage()
        QMessageBox.information(self, "完成", f"已生成 {len(generated)} 个文件")
    def on_print(self):
        files = self.get_files()
        if not files:
            QMessageBox.warning(self, "提示", "请先导入发票")
            return
        out_dir = self.line_out.text().strip() or None
        od = out_dir or os.path.dirname(files[0])
        target = os.path.join(od, "merged_2up.pdf")
        if not os.path.exists(target):
            QMessageBox.information(self, "提示", "未找到排版后的文件，请先排版")
            return
        copies = self.spin_copies.value()
        self.set_busy(True)
        self.statusBar().showMessage("正在打开打印对话框…")
        try:
            for _ in range(copies):
                try:
                    print_pdf(target)
                except Exception:
                    pass
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

def run_gui():
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
