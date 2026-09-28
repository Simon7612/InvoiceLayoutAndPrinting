"""readInvoice 读取/收集逻辑的单元测试。"""
import base64
import io

import pytest
from pypdf import PdfReader, PdfWriter

from readInvoice import (
    collect_pdfs,
    detect_ticket_document,
    read_document,
    read_pdf,
)


def make_pdf_bytes(width=200.0, height=100.0, text=None):
    """生成最小 PDF bytes；可选用标准 Helvetica 字体绘制一行文本。"""
    writer = PdfWriter()
    page = writer.add_blank_page(width=width, height=height)
    if text:
        from pypdf.generic import (
            DecodedStreamObject,
            DictionaryObject,
            NameObject,
        )

        stream = DecodedStreamObject()
        stream.set_data(
            f"BT /F1 14 Tf 20 100 Td ({text}) Tj ET".encode("ascii")
        )
        page[NameObject("/Contents")] = stream
        helv = DictionaryObject()
        helv[NameObject("/Type")] = NameObject("/Font")
        helv[NameObject("/Subtype")] = NameObject("/Type1")
        helv[NameObject("/BaseFont")] = NameObject("/Helvetica")
        fonts = DictionaryObject()
        fonts[NameObject("/F1")] = writer._add_object(helv)
        resources = DictionaryObject()
        resources[NameObject("/Font")] = fonts
        page[NameObject("/Resources")] = resources
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


# ---------------- collect_pdfs ----------------

def test_collect_pdfs_directory_filters_by_extension(tmp_path):
    (tmp_path / "b.pdf").write_bytes(b"x")
    (tmp_path / "a.ofd").write_bytes(b"x")
    (tmp_path / "c.xml").write_bytes(b"x")
    (tmp_path / "d.txt").write_bytes(b"x")
    (tmp_path / "e.PDF").write_bytes(b"x")  # 扩展名大小写不敏感
    assert collect_pdfs(str(tmp_path)) == [
        str(tmp_path / "a.ofd"),
        str(tmp_path / "b.pdf"),
        str(tmp_path / "c.xml"),
        str(tmp_path / "e.PDF"),
    ]


def test_collect_pdfs_single_file_returns_itself(tmp_path):
    f = tmp_path / "one.pdf"
    f.write_bytes(b"x")
    assert collect_pdfs(str(f)) == [str(f)]


# ---------------- read_pdf / read_document ----------------

def test_read_pdf_missing_file_raises():
    with pytest.raises(FileNotFoundError):
        read_pdf("Z:/no/such/file.pdf")


def test_read_pdf_rejects_non_pdf(tmp_path):
    f = tmp_path / "x.ofd"
    f.write_bytes(b"x")
    with pytest.raises(ValueError):
        read_pdf(str(f))


def test_read_document_with_pdf(tmp_path):
    f = tmp_path / "doc.pdf"
    f.write_bytes(make_pdf_bytes())
    reader = read_document(str(f))
    assert isinstance(reader, PdfReader)
    assert len(reader.pages) == 1


def test_read_document_xml_with_embedded_pdf(tmp_path):
    payload = base64.b64encode(make_pdf_bytes()).decode("ascii")
    f = tmp_path / "doc.xml"
    f.write_text(f"<invoice><pdf>{payload}</pdf></invoice>", encoding="ascii")
    reader = read_document(str(f))
    assert len(reader.pages) == 1


def test_read_document_xml_without_payload_raises(tmp_path):
    f = tmp_path / "empty.xml"
    f.write_text("<invoice><note>plain text only</note></invoice>", encoding="ascii")
    with pytest.raises(RuntimeError):
        read_document(str(f))


def test_read_document_xml_rejects_entity_expansion(tmp_path):
    # defusedxml 必须在解析阶段拒绝 DTD 实体（billion laughs 类攻击）
    from defusedxml.common import EntitiesForbidden

    bomb = (
        '<?xml version="1.0"?>\n'
        f'<!DOCTYPE invoice [<!ENTITY a "{"A" * 64}">]>\n'
        "<invoice><pdf>&a;&a;&a;&a;&a;&a;&a;&a;</pdf></invoice>"
    )
    f = tmp_path / "bomb.xml"
    f.write_text(bomb, encoding="ascii")
    with pytest.raises(EntitiesForbidden):
        read_document(str(f))


# ---------------- OFD 转换回退链 ----------------

def test_read_document_ofd_uses_ofdparser(monkeypatch, tmp_path):
    import readInvoice as ri

    class FakeParser:
        def __init__(self, b64):
            pass
        def ofd2pdf(self):
            return make_pdf_bytes()

    monkeypatch.setattr(ri, "_HAS_OFDPARSER", True)
    monkeypatch.setattr(ri, "_HAS_EASYOFD", False)
    monkeypatch.setattr(ri, "OfdParser", FakeParser, raising=False)
    f = tmp_path / "doc.ofd"
    f.write_bytes(b"PK\x03\x04fake-ofd")
    assert len(read_document(str(f)).pages) == 1


def test_read_document_ofd_falls_back_to_easyofd(monkeypatch, tmp_path):
    import readInvoice as ri

    class GarbageParser:
        def __init__(self, b64):
            pass
        def ofd2pdf(self):
            return b"short"  # 长度 <= 10 视为失败，触发回退

    class FakeEasyofd:
        @staticmethod
        def ofd2pdf(data):
            return make_pdf_bytes()

    monkeypatch.setattr(ri, "_HAS_OFDPARSER", True)
    monkeypatch.setattr(ri, "OfdParser", GarbageParser, raising=False)
    monkeypatch.setattr(ri, "_HAS_EASYOFD", True)
    monkeypatch.setattr(ri, "easyofd", FakeEasyofd, raising=False)
    f = tmp_path / "doc.ofd"
    f.write_bytes(b"PK\x03\x04fake-ofd")
    assert len(read_document(str(f)).pages) == 1


def test_read_document_ofd_without_libraries_raises(monkeypatch, tmp_path):
    import readInvoice as ri

    monkeypatch.setattr(ri, "_HAS_OFDPARSER", False)
    monkeypatch.setattr(ri, "_HAS_EASYOFD", False)
    f = tmp_path / "doc.ofd"
    f.write_bytes(b"PK\x03\x04fake-ofd")
    with pytest.raises(RuntimeError):
        read_document(str(f))


# ---------------- detect_ticket_document ----------------

def test_detect_ticket_matches_keywords_and_orientation():
    reader = PdfReader(io.BytesIO(make_pdf_bytes(width=300, height=200, text="Railway departure gate 12 seat")))
    matched, orient = detect_ticket_document(reader)
    assert matched is True
    assert orient == "landscape"  # 宽 >= 高


def test_detect_ticket_negative():
    reader = PdfReader(io.BytesIO(make_pdf_bytes(text="Invoice total amount due")))
    matched, _ = detect_ticket_document(reader)
    assert matched is False


def test_detect_ticket_single_english_keyword_not_matched():
    # 单个英文常见词不构成车票证据（需 ≥2 命中）
    reader = PdfReader(io.BytesIO(make_pdf_bytes(text="gate 12 information")))
    matched, _ = detect_ticket_document(reader)
    assert matched is False


def test_detect_ticket_no_text_page():
    reader = PdfReader(io.BytesIO(make_pdf_bytes(width=100, height=200)))
    matched, orient = detect_ticket_document(reader)
    assert matched is False
    assert orient == "portrait"  # 宽 < 高
