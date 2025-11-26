from typing import Optional, List
from copy import deepcopy
from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from pypdf import Transformation
from pypdf.generic import RectangleObject, DictionaryObject, NameObject, ArrayObject, FloatObject

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

def write_writer(writer: PdfWriter, output_path: str) -> None:
    with open(output_path, "wb") as f:
        writer.write(f)