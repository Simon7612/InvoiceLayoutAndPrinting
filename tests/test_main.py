"""main.py CLI 批处理流程的单元测试（不触发打印）。"""
import io

from pypdf import PdfReader, PdfWriter

import main as main_mod


def _single_page_pdf() -> bytes:
    writer = PdfWriter()
    writer.add_blank_page(width=200, height=100)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def test_process_batch_two_up(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    for i in range(2):
        (src / f"inv{i}.pdf").write_bytes(_single_page_pdf())
    out = tmp_path / "out"
    out.mkdir()

    main_mod.process(str(src), str(out), do_print=False)

    outputs = sorted(out.glob("*_2up.pdf"))
    assert [f.name for f in outputs] == ["inv0_2up.pdf", "inv1_2up.pdf"]
    for f in outputs:
        # 单页文件 2-up 后仍是 1 页
        assert len(PdfReader(str(f)).pages) == 1


def test_process_output_defaults_to_source_dir(tmp_path):
    f = tmp_path / "one.pdf"
    f.write_bytes(_single_page_pdf())

    main_mod.process(str(f), None, do_print=False)

    out = tmp_path / "one_2up.pdf"
    assert out.exists()
    assert len(PdfReader(str(out)).pages) == 1
