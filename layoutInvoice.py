from typing import Optional, List
import logging

from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from pypdf import Transformation
from pypdf.generic import (
    RectangleObject,
    NameObject,
    ArrayObject,
    FloatObject,
    DecodedStreamObject,
)

logger = logging.getLogger(__name__)

def _cropbox_metrics(p: PageObject) -> tuple[float, float, float, float]:
    cb = RectangleObject(p.cropbox)
    w = float(cb.width)
    h = float(cb.height)
    left = float(cb.left)
    bottom = float(cb.bottom)
    return w, h, left, bottom

def _adjust_merged_annots(dst: PageObject, n1: int, dx1: float, dy1: float, n2: int, dx2: float, dy2: float) -> None:
    ann = dst.get("/Annots")
    if not ann:
        return
    try:
        for idx, a in enumerate(ann):
            o = a.get_object()
            r = RectangleObject(o["/Rect"])
            llx, lly = r.lower_left
            urx, ury = r.upper_right
            if idx < n1:
                nx, ny = dx1, dy1
            else:
                nx, ny = dx2, dy2
            new_rect = ArrayObject([
                FloatObject(float(llx) + nx),
                FloatObject(float(lly) + ny),
                FloatObject(float(urx) + nx),
                FloatObject(float(ury) + ny),
            ])
            o[NameObject("/Rect")] = new_rect
    except Exception:
        # 印章错位是本项目的核心修复点，失败必须留下可排查的痕迹
        logger.warning("电子印章注释(/Annots)平移失败", exc_info=True)

def two_up_vertical(reader: PdfReader) -> PdfWriter:
    # 旧 CLI 入口：与 GUI 共用同一套 2-up 竖向合成实现
    return two_up_vertical_pages(reader.pages)

def two_up_vertical_pages(pages: List[PageObject]) -> PdfWriter:
    writer = PdfWriter()
    n = len(pages)
    for i in range(0, n, 2):
        p1 = pages[i]
        w1, h1, l1, b1 = _cropbox_metrics(p1)
        p2: Optional[PageObject] = pages[i + 1] if i + 1 < n else None
        if p2 is not None:
            w2, h2, l2, b2 = _cropbox_metrics(p2)
        else:
            w2, h2, l2, b2 = w1, h1, l1, b1
        blank_w = max(w1, w2)
        blank_h = h1 + h2
        blank = PageObject.create_blank_page(width=blank_w, height=blank_h)
        t1 = Transformation().translate(-l1, -b1 + (blank_h - h1))
        blank.merge_transformed_page(p1, t1)
        if p2 is not None:
            t2 = Transformation().translate(-l2, -b2)
            blank.merge_transformed_page(p2, t2)
        # merge_transformed_page 会原样复制 /Annots（电子印章），这里按与内容相同的
        # 平移量把两页注释的 /Rect 平移到合成后的位置
        writer.add_page(blank)
        page = writer.pages[-1]
        n1 = len(p1.get("/Annots") or [])
        n2 = len(p2.get("/Annots") or []) if p2 is not None else 0
        _adjust_merged_annots(
            page, n1, -l1, (blank_h - h1) - b1,
            n2, -l2 if p2 is not None else 0.0, -b2 if p2 is not None else 0.0,
        )
    return writer

def one_up_pages(pages: List[PageObject]) -> PdfWriter:
    """每页一张，按原始尺寸逐页输出。

    适用于“单页一张”的排版；后续可在调用层增加旋转/纸张方向控制。
    """
    writer = PdfWriter()
    for p in pages:
        writer.add_page(p)
    return writer

def two_up_horizontal_pages(pages: List[PageObject]) -> PdfWriter:
    """双页左右排版（2-up horizontal）。

    将相邻两页并排放置在同一页面的左、右。
    """
    writer = PdfWriter()
    n = len(pages)
    for i in range(0, n, 2):
        p1 = pages[i]
        w1, h1, l1, b1 = _cropbox_metrics(p1)
        p2: Optional[PageObject] = pages[i + 1] if i + 1 < n else None
        if p2 is not None:
            w2, h2, l2, b2 = _cropbox_metrics(p2)
        else:
            w2, h2, l2, b2 = w1, h1, l1, b1
        blank_w = w1 + w2
        blank_h = max(h1, h2)
        blank = PageObject.create_blank_page(width=blank_w, height=blank_h)
        # 左侧页：对齐底边
        t1 = Transformation().translate(-l1, -b1)
        blank.merge_transformed_page(p1, t1)
        # 右侧页：在左页宽度之后放置
        if p2 is not None:
            t2 = Transformation().translate(-l2 + w1, -b2)
            blank.merge_transformed_page(p2, t2)
        # merge_transformed_page 会原样复制 /Annots（电子印章），这里按与内容相同的
        # 平移量把两页注释的 /Rect 平移到合成后的位置
        writer.add_page(blank)
        page = writer.pages[-1]
        n1 = len(p1.get("/Annots") or [])
        n2 = len(p2.get("/Annots") or []) if p2 is not None else 0
        _adjust_merged_annots(
            page, n1, -l1, -b1,
            n2, (-l2 + w1) if p2 is not None else 0.0, -b2 if p2 is not None else 0.0,
        )
    return writer

def _transform_annots(dst: PageObject, start: int, count: int,
                      sx: float, sy: float, dx: float, dy: float) -> None:
    """对 dst 上 [start, start+count) 范围的注释矩形做仿射变换：x' = sx*x + dx。

    merge_transformed_page 复制注释时保留原始坐标，缩放类布局（宫格）
    需要在合并后按与内容相同的变换修正 /Rect。
    """
    ann = dst.get("/Annots")
    if not ann:
        return
    try:
        for a in list(ann)[start:start + count]:
            o = a.get_object()
            r = RectangleObject(o["/Rect"])
            llx, lly = r.lower_left
            urx, ury = r.upper_right
            o[NameObject("/Rect")] = ArrayObject([
                FloatObject(float(llx) * sx + dx),
                FloatObject(float(lly) * sy + dy),
                FloatObject(float(urx) * sx + dx),
                FloatObject(float(ury) * sy + dy),
            ])
    except Exception:
        # 印章错位是本项目的核心修复点，失败必须留下可排查的痕迹
        logger.warning("电子印章注释(/Annots)变换失败", exc_info=True)


def grid_pages(pages: List[PageObject], rows: int, cols: int) -> PdfWriter:
    """rows×cols 宫格排版：每张源页等比缩放到单元格内完整显示。

    单元格尺寸取每组首页；不足一组时重复最后一页填充。
    注释(/Annots，电子印章)随内容做相同的缩放平移。
    """
    if rows < 1 or cols < 1:
        raise ValueError("grid rows/cols must be >= 1")
    writer = PdfWriter()
    n = len(pages)
    per_sheet = rows * cols
    i = 0
    while i < n:
        w, h, _, _ = _cropbox_metrics(pages[i])
        sheet_w, sheet_h = w * cols, h * rows
        blank = PageObject.create_blank_page(width=sheet_w, height=sheet_h)
        group = [pages[j] if j < n else pages[n - 1] for j in range(i, i + per_sheet)]
        specs = []
        for idx, pg in enumerate(group):
            pw, ph, pl, pb = _cropbox_metrics(pg)
            row, col = divmod(idx, cols)
            s = min(w / pw, h / ph)
            cell_x = col * w
            cell_y = sheet_h - (row + 1) * h  # row 0 在最上面
            dx = cell_x - s * pl
            dy = cell_y - s * pb
            specs.append((s, dx, dy, len(pg.get("/Annots") or [])))
            blank.merge_transformed_page(pg, Transformation().scale(s, s).translate(dx, dy))
        writer.add_page(blank)
        page = writer.pages[-1]
        start = 0
        for s, dx, dy, cnt in specs:
            if cnt:
                _transform_annots(page, start, cnt, s, s, dx, dy)
            start += cnt
        i += per_sheet
    return writer


def four_up_grid_pages(pages: List[PageObject]) -> PdfWriter:
    """四宫格（2×2）排版，固定为横向布局。"""
    return grid_pages(pages, 2, 2)

# --- 切割线与布局调度 ---

def _append_content_stream(page: PageObject, data: bytes) -> None:
    new_stream = DecodedStreamObject()
    new_stream.set_data(data)
    existing = page.get("/Contents")
    if existing:
        if isinstance(existing, ArrayObject):
            existing.append(new_stream)
        else:
            page[NameObject("/Contents")] = ArrayObject([existing, new_stream])
    else:
        page[NameObject("/Contents")] = new_stream

def _draw_cut_lines(page: PageObject, vertical_fracs: List[float], horizontal_fracs: List[float]) -> None:
    """按宽度/高度比例绘制竖线与横线切割线。"""
    cb = RectangleObject(page.cropbox)
    w = float(cb.width)
    h = float(cb.height)
    margin = 6.0
    parts: list[str] = [
        "0 0 0 RG",   # stroke color: black
        "0.8 w",      # line width
    ]
    for f in vertical_fracs:
        x = w * f
        parts += [f"{x:.2f} {margin:.2f} m", f"{x:.2f} {h - margin:.2f} l", "S"]
    for f in horizontal_fracs:
        y = h * f
        parts += [f"{margin:.2f} {y:.2f} m", f"{w - margin:.2f} {y:.2f} l", "S"]
    content = ("\n".join(parts) + "\n").encode("ascii")
    _append_content_stream(page, content)

class LayoutMode:
    ONE_UP = "one_up"
    TWO_UP_VERTICAL = "two_up_vertical"
    TWO_UP_HORIZONTAL = "two_up_horizontal"
    FOUR_UP = "four_up"
    CUSTOM_GRID = "custom_grid"

class Orientation:
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"

def _rotate_pages_for_orientation(pages: List[PageObject], orientation: Optional[str]) -> List[PageObject]:
    if orientation is None:
        # None 表示不做方向预处理（CLI 批处理保持旧的合成结果）
        return list(pages)
    out: List[PageObject] = []
    for p in pages:
        w, h, _left, _bottom = _cropbox_metrics(p)
        if orientation == Orientation.LANDSCAPE and h > w:
            try:
                p = p.rotate(90)  # pypdf 支持 rotate/rotate_clockwise
            except Exception:
                pass
        elif orientation == Orientation.PORTRAIT and w > h:
            try:
                p = p.rotate(90)
            except Exception:
                pass
        out.append(p)
    return out

def compose_pages(pages: List[PageObject], layout_mode: str, orientation: Optional[str],
                  add_cutlines: bool, grid: Optional[List[int]] = None) -> PdfWriter:
    # 方向预处理：仅对单页/重复场景有意义；对合成页尺寸影响有限，尽量保持输入页方向一致
    pages2 = _rotate_pages_for_orientation(pages, orientation)
    if layout_mode == LayoutMode.ONE_UP:
        writer = one_up_pages(pages2)
        cut = ([], [])
    elif layout_mode == LayoutMode.TWO_UP_VERTICAL:
        writer = two_up_vertical_pages(pages2)
        cut = ([], [0.5])
    elif layout_mode == LayoutMode.TWO_UP_HORIZONTAL:
        writer = two_up_horizontal_pages(pages2)
        cut = ([0.5], [])
    elif layout_mode == LayoutMode.FOUR_UP:
        writer = four_up_grid_pages(pages2)
        cut = ([0.5], [0.5])
    elif layout_mode == LayoutMode.CUSTOM_GRID:
        if not grid or len(grid) != 2 or int(grid[0]) < 1 or int(grid[1]) < 1:
            raise ValueError("custom grid layout requires (rows, cols)")
        rows, cols = int(grid[0]), int(grid[1])
        writer = grid_pages(pages2, rows, cols)
        cut = ([c / cols for c in range(1, cols)], [r / rows for r in range(1, rows)])
    else:
        raise ValueError(f"unknown layout mode: {layout_mode}")
    if add_cutlines:
        for pg in writer.pages:
            _draw_cut_lines(pg, *cut)
    return writer

def write_writer(writer: PdfWriter, output_path: str) -> None:
    with open(output_path, "wb") as f:
        writer.write(f)
