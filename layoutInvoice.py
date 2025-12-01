from typing import Optional, List
from copy import deepcopy
from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from pypdf import Transformation
from pypdf.generic import (
    RectangleObject,
    DictionaryObject,
    NameObject,
    ArrayObject,
    FloatObject,
    DecodedStreamObject,
)

def _cropbox_metrics(p: PageObject) -> tuple[float, float, float, float]:
    cb = RectangleObject(p.cropbox)
    w = float(cb.width)
    h = float(cb.height)
    left = float(cb.left)
    bottom = float(cb.bottom)
    return w, h, left, bottom

def _move_annots(src: PageObject, dst: PageObject, dx: float, dy: float) -> None:
    try:
        annots = src.get("/Annots")
        if not annots:
            return
        for a in annots:
            o = a.get_object()
            r = RectangleObject(o["/Rect"])
            llx, lly = r.lower_left
            urx, ury = r.upper_right
            new_rect = ArrayObject([
                FloatObject(float(llx) + dx),
                FloatObject(float(lly) + dy),
                FloatObject(float(urx) + dx),
                FloatObject(float(ury) + dy),
            ])
            new = DictionaryObject()
            for k in o.keys():
                if k == NameObject("/Rect"):
                    new[k] = new_rect
                else:
                    new[k] = o[k]
            dst.add_annotation(new)
    except Exception:
        pass

def _adjust_merged_annots(dst: PageObject, n1: int, dx1: float, dy1: float, n2: int, dx2: float, dy2: float) -> None:
    ann = dst.get("/Annots")
    if not ann:
        return
    try:
        total = len(ann)
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
        pass

def two_up_vertical(reader: PdfReader) -> PdfWriter:
    writer = PdfWriter()
    pages = reader.pages
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
        _move_annots(p1, blank, -l1, (blank_h - h1))
        writer.add_page(blank)
        page = writer.pages[-1]
        n1 = len(p1.get("/Annots") or [])
        n2 = len(p2.get("/Annots") or []) if p2 is not None else 0
        _adjust_merged_annots(page, n1, -l1, (blank_h - h1), n2, -l2 if p2 is not None else 0.0, 0.0)
        
    return writer

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
        _move_annots(p1, blank, -l1, (blank_h - h1))
        writer.add_page(blank)
        page = writer.pages[-1]
        n1 = len(p1.get("/Annots") or [])
        n2 = len(p2.get("/Annots") or []) if p2 is not None else 0
        _adjust_merged_annots(page, n1, -l1, (blank_h - h1), n2, -l2 if p2 is not None else 0.0, 0.0)
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
        _move_annots(p1, blank, -l1, -b1)
        writer.add_page(blank)
        page = writer.pages[-1]
        n1 = len(p1.get("/Annots") or [])
        n2 = len(p2.get("/Annots") or []) if p2 is not None else 0
        _adjust_merged_annots(page, n1, -l1, -b1, n2, -l2 if p2 is not None else 0.0, -b2 if p2 is not None else 0.0)
    return writer

def four_up_grid_pages(pages: List[PageObject]) -> PdfWriter:
    """四宫格（2×2）排版，固定为横向布局。

    简化实现：按首页尺寸创建目标页，并对每个源页进行 0.5 缩放后分别放置到四个象限。
    注：此实现未缩放注释；若需支持注释缩放，需对 /Annots 中的矩形坐标按比例调整。
    """
    writer = PdfWriter()
    n = len(pages)
    i = 0
    while i < n:
        # 使用第一个页面的尺寸作为基准
        p1 = pages[i]
        w, h, l, b = _cropbox_metrics(p1)
        blank_w = w * 2
        blank_h = h * 2
        blank = PageObject.create_blank_page(width=blank_w, height=blank_h)
        # 放置四个页面，若不足四页则重复最后一页
        group = [pages[j] if j < n else pages[n - 1] for j in [i, i + 1, i + 2, i + 3]]
        positions = [
            (0.0, h),        # 左上
            (w, h),          # 右上
            (0.0, 0.0),      # 左下
            (w, 0.0),        # 右下
        ]
        for idx, (pg, (tx, ty)) in enumerate(zip(group, positions)):
            pw, ph, pl, pb = _cropbox_metrics(pg)
            # 统一按 0.5 缩放；偏移到对应象限
            t = Transformation().scale(0.5, 0.5).translate(-pl + tx, -pb + ty)
            blank.merge_transformed_page(pg, t)
        writer.add_page(blank)
        i += 4
    return writer

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

def _draw_center_lines(page: PageObject, vertical: bool = False, horizontal: bool = False) -> None:
    cb = RectangleObject(page.cropbox)
    w = float(cb.width)
    h = float(cb.height)
    margin = 6.0
    parts: list[str] = [
        "0 0 0 RG",   # stroke color: black
        "0.8 w",      # line width
    ]
    if vertical:
        x = w / 2.0
        parts += [f"{x:.2f} {margin:.2f} m", f"{x:.2f} {h - margin:.2f} l", "S"]
    if horizontal:
        y = h / 2.0
        parts += [f"{margin:.2f} {y:.2f} m", f"{w - margin:.2f} {y:.2f} l", "S"]
    content = ("\n".join(parts) + "\n").encode("ascii")
    _append_content_stream(page, content)

class LayoutMode:
    ONE_UP = "one_up"
    TWO_UP_VERTICAL = "two_up_vertical"
    TWO_UP_HORIZONTAL = "two_up_horizontal"
    FOUR_UP = "four_up"

class Orientation:
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"

def _rotate_pages_for_orientation(pages: List[PageObject], orientation: str) -> List[PageObject]:
    out: List[PageObject] = []
    for p in pages:
        w, h, l, b = _cropbox_metrics(p)
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

def compose_pages(pages: List[PageObject], layout_mode: str, orientation: str, add_cutlines: bool) -> PdfWriter:
    # 方向预处理：仅对单页/重复场景有意义；对合成页尺寸影响有限，尽量保持输入页方向一致
    pages2 = _rotate_pages_for_orientation(pages, orientation)
    if layout_mode == LayoutMode.ONE_UP:
        writer = one_up_pages(pages2)
    elif layout_mode == LayoutMode.TWO_UP_VERTICAL:
        writer = two_up_vertical_pages(pages2)
    elif layout_mode == LayoutMode.TWO_UP_HORIZONTAL:
        writer = two_up_horizontal_pages(pages2)
    elif layout_mode == LayoutMode.FOUR_UP:
        writer = four_up_grid_pages(pages2)
    else:
        raise ValueError(f"unknown layout mode: {layout_mode}")
    if add_cutlines:
        for pg in writer.pages:
            if layout_mode == LayoutMode.TWO_UP_VERTICAL:
                _draw_center_lines(pg, horizontal=True)
            elif layout_mode == LayoutMode.TWO_UP_HORIZONTAL:
                _draw_center_lines(pg, vertical=True)
            elif layout_mode == LayoutMode.FOUR_UP:
                _draw_center_lines(pg, vertical=True, horizontal=True)
    return writer

def write_writer(writer: PdfWriter, output_path: str) -> None:
    with open(output_path, "wb") as f:
        writer.write(f)
