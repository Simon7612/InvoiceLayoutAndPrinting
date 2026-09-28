"""CLI 与 GUI 共用的排版打印流水线。

「收集 → 读取/转换 → 车票识别 → 合成 → 写出 → 打印」的编排只在这里定义一次，
main.py 的批处理与 gui 的后台线程都是它的调用方。本模块不依赖 Qt，
便于无 GUI 环境下测试。
"""
import logging
import os
from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple

from pypdf import PdfReader

from readInvoice import collect_pdfs, detect_ticket_document, read_document, read_pdf
from layoutInvoice import LayoutMode, compose_pages, write_writer
from printInvoice import print_pdf

logger = logging.getLogger(__name__)

# 各布局模式对应的输出文件名（排版与打印按钮共用）
LAYOUT_OUTPUT_NAMES = {
    LayoutMode.ONE_UP: "merged_1up.pdf",
    LayoutMode.TWO_UP_VERTICAL: "merged_2up_v.pdf",
    LayoutMode.TWO_UP_HORIZONTAL: "merged_2up_h.pdf",
    LayoutMode.FOUR_UP: "merged_4up.pdf",
}

# load_documents 的进度回调：(当前序号, 总数, 文件名)
LoadProgress = Callable[[int, int, str], None]
# run_compose_job 的进度回调：一句话状态
MessageProgress = Callable[[str], None]
CancelCheck = Callable[[], bool]


def output_name(mode: str, grid: Optional[List[int]] = None) -> str:
    """布局模式对应的输出文件名；custom_grid 需提供 [rows, cols]。"""
    if mode == LayoutMode.CUSTOM_GRID:
        if not grid or len(grid) != 2:
            raise ValueError("custom_grid 输出名需要 grid=[rows, cols]")
        return f"merged_grid{int(grid[0])}x{int(grid[1])}.pdf"
    return LAYOUT_OUTPUT_NAMES[mode]


@dataclass
class ComposeJob:
    """一次「合成 → 写出 → 打印」任务的全部参数。"""

    pages: List                        # 源页列表（PageObject）
    mode: str                          # LayoutMode.*
    orientation: Optional[str] = None  # Orientation.*；None 表示不做方向预处理（CLI 旧行为）
    add_cutlines: bool = False
    grid: Optional[List[int]] = None   # custom_grid 模式的 [rows, cols]
    out_path: str = ""
    do_print: bool = False
    copies: int = 1
    printer: Optional[str] = None


def run_compose_job(job: ComposeJob, progress: Optional[MessageProgress] = None,
                    should_cancel: Optional[CancelCheck] = None) -> str:
    """执行任务：合成 → 写出 →（可选）逐份打印，返回输出路径。

    逐份打印失败只记日志不中断：产物已落盘，一份打印失败不应推翻整个任务。
    """
    def _notify(msg: str) -> None:
        if progress:
            progress(msg)

    _notify("正在排版合成…")
    writer = compose_pages(job.pages, job.mode, job.orientation,
                           job.add_cutlines, grid=job.grid)
    _notify("正在写出文件…")
    write_writer(writer, job.out_path)
    if job.do_print:
        for _ in range(job.copies):
            if should_cancel and should_cancel():
                break
            try:
                print_pdf(job.out_path, job.printer)
            except Exception as e:
                logger.warning("打印失败 %s: %s", job.out_path, e)
    return job.out_path


def read_document_lenient(path: str) -> PdfReader:
    """读取文档；统一入口 read_document 失败时回退按普通 PDF 读取。"""
    try:
        return read_document(path)
    except Exception:
        return read_pdf(path)


def load_documents(files: List[str], progress: Optional[LoadProgress] = None,
                   should_cancel: Optional[CancelCheck] = None,
                   ) -> Tuple[List[Tuple[str, PdfReader]], bool, str]:
    """逐个读取/转换文件（OFD/XML 转换较慢，调用方应放在后台线程）。

    返回 ([(源路径, PdfReader), ...], 是否发现车票, 建议方向)。
    多个文件的方向建议不一致时保留第一个。
    """
    readers: List[Tuple[str, PdfReader]] = []
    any_ticket = False
    suggested = ""
    total = len(files)
    for i, src in enumerate(files, 1):
        if should_cancel and should_cancel():
            break
        if progress:
            progress(i, total, os.path.basename(src))
        r = read_document_lenient(src)
        try:
            is_ticket, orient_hint = detect_ticket_document(r)
            if is_ticket:
                any_ticket = True
                if not suggested and orient_hint:
                    suggested = orient_hint
        except Exception:
            pass
        readers.append((src, r))
    return readers, any_ticket, suggested


def process(input_path: str, output_dir: Optional[str], do_print: bool) -> None:
    """CLI 批处理：逐个文件 2-up 竖向合成，保持旧的逐文件命名 <名>_2up.pdf。"""
    files = collect_pdfs(input_path)
    for src in files:
        reader = read_document(src)
        out_dir = output_dir or os.path.dirname(src)
        stem = os.path.splitext(os.path.basename(src))[0]
        job = ComposeJob(
            pages=list(reader.pages),
            mode=LayoutMode.TWO_UP_VERTICAL,
            out_path=os.path.join(out_dir, f"{stem}_2up.pdf"),
            do_print=do_print,
        )
        run_compose_job(job)
        logger.info("已生成 %s", job.out_path)
