"""layoutInvoice 合成逻辑的单元测试。

所有页面均用 pypdf 程序化生成，不依赖仓库中的真实发票 PDF。
注释（/Annots，即电子印章）平移是本项目修复过的核心行为，相关用例
锁定的期望值即“印章与页面内容保持相对静止”。
"""
import io

import pytest
from pypdf import PdfReader, PdfWriter
from pypdf._page import PageObject
from pypdf.generic import (
    ArrayObject,
    ContentStream,
    DecodedStreamObject,
    DictionaryObject,
    FloatObject,
    NameObject,
    RectangleObject,
)

from layoutInvoice import (
    LayoutMode,
    Orientation,
    compose_pages,
    four_up_grid_pages,
    grid_pages,
    one_up_pages,
    two_up_horizontal_pages,
    two_up_vertical,
    two_up_vertical_pages,
    write_writer,
)


def make_page(w, h, annot_rect=None, left=0.0, bottom=0.0):
    """生成 w×h 的页面；cropbox 原点可偏移，附带最小内容流与可选注释。"""
    page = PageObject.create_blank_page(width=left + w, height=bottom + h)
    page.cropbox = RectangleObject([left, bottom, left + w, bottom + h])
    stream = DecodedStreamObject()
    stream.set_data(b"0 0 1 RG 2 2 10 10 re S\n")
    page[NameObject("/Contents")] = stream
    if annot_rect is not None:
        annot = DictionaryObject()
        annot[NameObject("/Type")] = NameObject("/Annot")
        annot[NameObject("/Subtype")] = NameObject("/Square")
        annot[NameObject("/Rect")] = ArrayObject(
            [FloatObject(float(x)) for x in annot_rect]
        )
        page[NameObject("/Annots")] = ArrayObject([annot])
    return page


def make_reader(pages):
    writer = PdfWriter()
    for p in pages:
        writer.add_page(p)
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return PdfReader(buf)


def annot_rects(page):
    return [
        tuple(round(float(x), 1) for x in RectangleObject(a.get_object()["/Rect"]))
        for a in (page.get("/Annots") or [])
    ]


def cropbox_size(page):
    cb = RectangleObject(page.cropbox)
    return (round(float(cb.width), 1), round(float(cb.height), 1))


# ---------------- one_up ----------------

def test_one_up_preserves_order_and_size():
    pages = [make_page(200, 100), make_page(180, 90), make_page(300, 50)]
    out = one_up_pages(pages)
    assert len(out.pages) == 3
    assert [cropbox_size(p) for p in out.pages] == [(200.0, 100.0), (180.0, 90.0), (300.0, 50.0)]


# ---------------- two_up_vertical ----------------

def test_two_up_vertical_merges_dimensions():
    out = two_up_vertical_pages([make_page(200, 100), make_page(180, 90)])
    assert len(out.pages) == 1
    assert cropbox_size(out.pages[0]) == (200.0, 190.0)  # 宽取最大，高为两页之和


def test_two_up_vertical_page_count_and_odd_tail():
    out = two_up_vertical_pages(
        [make_page(200, 100), make_page(200, 100), make_page(200, 80)]
    )
    assert len(out.pages) == 2
    # 奇数页时末组只有上页：当前实现输出 2*h1 高、下半空白（锁定现状）
    assert cropbox_size(out.pages[1]) == (200.0, 160.0)


def test_two_up_vertical_moves_annots_to_each_half():
    p1 = make_page(200, 100, annot_rect=[10, 80, 60, 95])  # 页内靠上（票头处）
    p2 = make_page(200, 100, annot_rect=[10, 5, 60, 20])
    out = two_up_vertical_pages([p1, p2])
    assert annot_rects(out.pages[0]) == [(10.0, 180.0, 60.0, 195.0), (10.0, 5.0, 60.0, 20.0)]


def test_two_up_vertical_annots_with_offset_cropbox():
    # cropbox 原点 (10,20)：注释绝对坐标 [20,100,70,115] == 页内相对 [10,80,60,95]；
    # p2 原点 (5,3)：绝对 [15,8,65,23] == 页内相对 [10,5,60,20]
    p1 = make_page(200, 100, annot_rect=[20, 100, 70, 115], left=10, bottom=20)
    p2 = make_page(200, 100, annot_rect=[15, 8, 65, 23], left=5, bottom=3)
    out = two_up_vertical_pages([p1, p2])
    # 期望与零原点情形完全一致（印章相对页面内容的位置不变）
    assert annot_rects(out.pages[0]) == [(10.0, 180.0, 60.0, 195.0), (10.0, 5.0, 60.0, 20.0)]


def test_two_up_vertical_legacy_reader_entry():
    reader = make_reader([make_page(200, 100), make_page(180, 90)])
    legacy = two_up_vertical(reader)
    direct = two_up_vertical_pages(list(reader.pages))
    assert len(legacy.pages) == 1
    assert cropbox_size(legacy.pages[0]) == cropbox_size(direct.pages[0])


# ---------------- two_up_horizontal ----------------

def test_two_up_horizontal_merges_dimensions():
    out = two_up_horizontal_pages([make_page(200, 100), make_page(180, 90)])
    assert len(out.pages) == 1
    assert cropbox_size(out.pages[0]) == (380.0, 100.0)  # 宽为两页之和，高取最大


def test_two_up_horizontal_moves_annots_to_each_half():
    p1 = make_page(200, 100, annot_rect=[150, 10, 195, 30])
    p2 = make_page(180, 90, annot_rect=[10, 5, 60, 20])
    out = two_up_horizontal_pages([p1, p2])
    # 右页印章需随内容右移一个左页宽度（w1=200）
    assert annot_rects(out.pages[0]) == [(150.0, 10.0, 195.0, 30.0), (210.0, 5.0, 260.0, 20.0)]


# ---------------- four_up ----------------

@pytest.mark.parametrize("n_input, n_output", [(4, 1), (5, 2), (6, 2), (3, 1), (1, 1)])
def test_four_up_grid_page_count(n_input, n_output):
    out = four_up_grid_pages([make_page(100, 50) for _ in range(n_input)])
    assert len(out.pages) == n_output


def test_four_up_grid_dimensions():
    out = four_up_grid_pages([make_page(100, 50) for _ in range(4)])
    assert cropbox_size(out.pages[0]) == (200.0, 100.0)


def test_four_up_grid_transforms_annots_to_quadrants():
    # 同尺寸页面在 2×2 网格中不缩放（填满单元格），印章随内容平移到对应象限
    pages = [make_page(100, 50, annot_rect=[10, 10, 20, 20]) for _ in range(4)]
    out = four_up_grid_pages(pages)
    assert annot_rects(out.pages[0]) == [
        (10.0, 60.0, 20.0, 70.0),    # 左上（row0, col0）
        (110.0, 60.0, 120.0, 70.0),  # 右上
        (10.0, 10.0, 20.0, 20.0),    # 左下
        (110.0, 10.0, 120.0, 20.0),  # 右下
    ]


def test_grid_scales_oversized_pages_to_fit_cell():
    # 单元格取首页尺寸（100×50）；第二页 200×100 缩放 0.5 适配
    pages = [
        make_page(100, 50, annot_rect=[10, 5, 60, 20]),
        make_page(200, 100, annot_rect=[10, 10, 20, 20]),
    ]
    out = grid_pages(pages, 1, 2)
    assert cropbox_size(out.pages[0]) == (200.0, 50.0)
    assert annot_rects(out.pages[0]) == [(10.0, 5.0, 60.0, 20.0), (105.0, 5.0, 110.0, 10.0)]


# ---------------- 通用网格 ----------------

def test_grid_one_by_two_keeps_scale_one():
    pages = [make_page(100, 50, annot_rect=[10, 5, 60, 20]) for _ in range(2)]
    out = grid_pages(pages, 1, 2)
    assert len(out.pages) == 1
    assert cropbox_size(out.pages[0]) == (200.0, 50.0)
    assert annot_rects(out.pages[0]) == [(10.0, 5.0, 60.0, 20.0), (110.0, 5.0, 160.0, 20.0)]


@pytest.mark.parametrize("n_input, rows, cols, n_output", [(7, 3, 2, 2), (13, 2, 3, 3), (1, 2, 2, 1)])
def test_grid_page_count_and_fill(n_input, rows, cols, n_output):
    out = grid_pages([make_page(100, 50) for _ in range(n_input)], rows, cols)
    assert len(out.pages) == n_output


def test_grid_rejects_invalid_shape():
    with pytest.raises(ValueError):
        grid_pages([make_page(100, 50)], 0, 2)


def test_compose_custom_grid_requires_valid_grid():
    pages = [make_page(100, 50)]
    with pytest.raises(ValueError):
        compose_pages(pages, LayoutMode.CUSTOM_GRID, Orientation.PORTRAIT, False)
    with pytest.raises(ValueError):
        compose_pages(pages, LayoutMode.CUSTOM_GRID, Orientation.PORTRAIT, False, grid=[0, 2])


def test_compose_custom_grid_dimensions_and_cutlines(tmp_path):
    pages = [make_page(100, 50) for _ in range(6)]
    out = compose_pages(pages, LayoutMode.CUSTOM_GRID, Orientation.PORTRAIT, True, grid=[2, 3])
    assert len(out.pages) == 1
    assert cropbox_size(out.pages[0]) == (300.0, 100.0)
    path = tmp_path / "grid.pdf"
    write_writer(out, str(path))
    reader = PdfReader(str(path))
    data = ContentStream(reader.pages[0]["/Contents"], reader).get_data()
    # 2 行 3 列：内部竖线 2 条 + 横线 1 条
    assert data.count(b" l\n") == 3
    assert b"0 0 0 RG" in data


# ---------------- compose_pages 调度 ----------------

def test_compose_unknown_mode_raises():
    with pytest.raises(ValueError):
        compose_pages([make_page(100, 50)], "no_such_mode", Orientation.PORTRAIT, False)


def test_compose_cutlines_add_drawing_ops(tmp_path):
    pages = [make_page(200, 100), make_page(180, 90)]

    out_with = compose_pages(pages, LayoutMode.TWO_UP_VERTICAL, Orientation.PORTRAIT, True)
    f1 = tmp_path / "with.pdf"
    write_writer(out_with, str(f1))
    reader = PdfReader(str(f1))
    data = ContentStream(reader.pages[0]["/Contents"], reader).get_data()
    assert b"0 0 0 RG" in data  # 切割线的黑色描边
    assert b"0 0 1 RG" in data  # 原页面内容仍在（源页的蓝色描边）

    out_without = compose_pages(pages, LayoutMode.TWO_UP_VERTICAL, Orientation.PORTRAIT, False)
    f2 = tmp_path / "without.pdf"
    write_writer(out_without, str(f2))
    reader2 = PdfReader(str(f2))
    data2 = ContentStream(reader2.pages[0]["/Contents"], reader2).get_data()
    assert b"0 0 0 RG" not in data2


def test_compose_orientation_rotates_mismatched_pages():
    # 注意：pypdf 的 rotate() 会原地修改页面对象，故每次合成使用新页面
    out = compose_pages([make_page(100, 200)], LayoutMode.ONE_UP, Orientation.LANDSCAPE, False)
    assert int(out.pages[0]["/Rotate"]) == 90

    out = compose_pages([make_page(200, 100)], LayoutMode.ONE_UP, Orientation.PORTRAIT, False)
    assert int(out.pages[0]["/Rotate"]) == 90

    out = compose_pages([make_page(100, 200)], LayoutMode.ONE_UP, Orientation.PORTRAIT, False)
    assert "/Rotate" not in out.pages[0]


def test_write_writer_roundtrip(tmp_path):
    out = two_up_vertical_pages([make_page(200, 100), make_page(180, 90)])
    path = tmp_path / "merged_2up_v.pdf"
    write_writer(out, str(path))
    reader = PdfReader(str(path))
    assert len(reader.pages) == 1
    assert cropbox_size(reader.pages[0]) == (200.0, 190.0)
