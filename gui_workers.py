"""GUI 后台线程：耗时操作交给 pipeline 执行，避免阻塞 UI。

线程本身只做信号转发与取消检查；读取/车票识别/合成/打印的编排都在 pipeline。
"""
from typing import List

from PyQt6.QtCore import QThread, pyqtSignal

import pipeline


class LoadWorker(QThread):
    """后台线程：读取/转换文件（OFD/XML 转换较慢）并识别车票。"""

    progress = pyqtSignal(int, int, str)      # (当前序号, 总数, 文件名)
    ticket_detected = pyqtSignal(bool, str)   # (是否车票, 建议方向)
    loaded = pyqtSignal(list)                 # [(源路径, PdfReader), ...]
    failed = pyqtSignal(str)

    def __init__(self, files: List[str], parent=None):
        super().__init__(parent)
        self._files = files

    def run(self):
        try:
            readers, any_ticket, suggested = pipeline.load_documents(
                self._files,
                progress=lambda i, total, name: self.progress.emit(i, total, name),
                should_cancel=self.isInterruptionRequested,
            )
            if self.isInterruptionRequested():
                return
            self.ticket_detected.emit(any_ticket, suggested)
            self.loaded.emit(readers)
        except Exception as e:
            self.failed.emit(str(e))


class ComposeWorker(QThread):
    """后台线程：页面合成、写出与打印。"""

    progress = pyqtSignal(str)
    composed = pyqtSignal(str)  # 输出文件路径
    failed = pyqtSignal(str)

    def __init__(self, job: pipeline.ComposeJob, parent=None):
        super().__init__(parent)
        self._job = job

    def run(self):
        try:
            pipeline.run_compose_job(
                self._job,
                progress=self.progress.emit,
                should_cancel=self.isInterruptionRequested,
            )
            self.composed.emit(self._job.out_path)
        except Exception as e:
            self.failed.emit(str(e))
