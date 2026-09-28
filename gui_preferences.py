"""QSettings 偏好持久化：窗口尺寸与排版/打印设置。

注意：程序化 setChecked 不触发 idClicked，restore() 末尾必须手动同步
派生 UI 状态（双页卡片副标题、网格可见性）。
"""
from PyQt6.QtCore import QSettings

SETTINGS_ORG = "Simon Chan"
SETTINGS_APP = "InvoiceLayoutAndPrinting"


class Preferences:
    """MainWindow 各项偏好的保存与恢复，读写键与 GUI 布局控件一一对应。"""

    def __init__(self, win):
        self._win = win

    def restore(self) -> None:
        w = self._win
        s = QSettings(SETTINGS_ORG, SETTINGS_APP)
        try:
            geo = s.value("window/geometry")
            if geo is not None:
                w.restoreGeometry(geo)
        except Exception:
            pass
        try:
            tile = s.value("layout/tile", 1, type=int)
            if tile == 0:
                w.btn_layout_custom.setChecked(True)
            elif tile == 1:
                w.btn_layout_one.setChecked(True)
            elif tile == 2:
                w.btn_layout_two_v.setChecked(True)
            elif tile == 3:
                w.btn_layout_four.setChecked(True)
            if s.value("layout/portrait", True, type=bool):
                w.btn_portrait.setChecked(True)
            else:
                w.btn_landscape.setChecked(True)
            w.chk_cutline.setChecked(s.value("layout/cutlines", False, type=bool))
            w.spin_grid_rows.setValue(s.value("layout/grid_rows", 2, type=int))
            w.spin_grid_cols.setValue(s.value("layout/grid_cols", 2, type=int))
            w.spin_copies.setValue(s.value("print/copies", 1, type=int))
            w.chk_print.setChecked(s.value("print/after_layout", False, type=bool))
            w.line_out.setText(s.value("paths/output_dir", "", type=str))
            idx = s.value("print/printer_index", 0, type=int)
            if 0 <= idx < w.combo_printer.count():
                w.combo_printer.setCurrentIndex(idx)
        except Exception:
            pass
        # 同步派生 UI 状态（程序化 setChecked 不触发 idClicked）
        try:
            w.update_two_tile_caption()
            w.update_grid_visibility()
        except Exception:
            pass

    def save(self) -> None:
        w = self._win
        s = QSettings(SETTINGS_ORG, SETTINGS_APP)
        s.setValue("window/geometry", w.saveGeometry())
        s.setValue("layout/tile", w.group_layout_tiles.checkedId())
        s.setValue("layout/portrait", w.btn_portrait.isChecked())
        s.setValue("layout/cutlines", w.chk_cutline.isChecked())
        s.setValue("layout/grid_rows", w.spin_grid_rows.value())
        s.setValue("layout/grid_cols", w.spin_grid_cols.value())
        s.setValue("print/copies", w.spin_copies.value())
        s.setValue("print/after_layout", w.chk_print.isChecked())
        s.setValue("print/printer_index", w.combo_printer.currentIndex())
        s.setValue("paths/output_dir", w.line_out.text().strip())
