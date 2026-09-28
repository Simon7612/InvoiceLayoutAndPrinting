"""可复用的 GUI 小部件。"""
from typing import Callable, List

from PyQt6.QtWidgets import QWidget


class ImportDropArea(QWidget):
    """接收文件拖拽的区域；拖入后通过 on_dropped 回调交出本地路径列表。"""

    def __init__(self, on_dropped: Callable[[List[str]], None]):
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
