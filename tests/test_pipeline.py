"""pipeline 流水线的单元测试（不依赖 Qt，不触发真实打印）。"""
import io

import pytest
from pypdf import PdfReader, PdfWriter

import pipeline
from layoutInvoice import LayoutMode, Orientation
from pipeline import ComposeJob, load_documents, output_name, run_compose_job


def _single_page_pdf_bytes() -> bytes:
    writer = PdfWriter()
    writer.add_blank_page(width=200, height=100)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def _pages(n: int) -> list:
    reader = PdfReader(io.BytesIO(_single_page_pdf_bytes()))
    return list(reader.pages) * n


# ---------------- output_name ----------------

def test_output_name_all_modes():
    assert output_name(LayoutMode.ONE_UP) == "merged_1up.pdf"
    assert output_name(LayoutMode.TWO_UP_VERTICAL) == "merged_2up_v.pdf"
    assert output_name(LayoutMode.TWO_UP_HORIZONTAL) == "merged_2up_h.pdf"
    assert output_name(LayoutMode.FOUR_UP) == "merged_4up.pdf"


def test_output_name_custom_grid():
    assert output_name(LayoutMode.CUSTOM_GRID, [2, 3]) == "merged_grid2x3.pdf"


def test_output_name_custom_grid_requires_grid():
    with pytest.raises(ValueError):
        output_name(LayoutMode.CUSTOM_GRID)
    with pytest.raises(ValueError):
        output_name(LayoutMode.CUSTOM_GRID, [2])


# ---------------- run_compose_job ----------------

def test_run_compose_job_writes_file_and_reports_progress(tmp_path):
    out = tmp_path / "merged_2up_v.pdf"
    messages = []
    job = ComposeJob(
        pages=_pages(2),
        mode=LayoutMode.TWO_UP_VERTICAL,
        orientation=Orientation.PORTRAIT,
        out_path=str(out),
    )
    result = run_compose_job(job, progress=messages.append)

    assert result == str(out)
    assert out.exists()
    assert len(PdfReader(str(out)).pages) == 1  # 两个单页源页合成 1 页
    assert messages == ["正在排版合成…", "正在写出文件…"]


def test_run_compose_job_prints_requested_copies(tmp_path, monkeypatch):
    out = tmp_path / "merged_1up.pdf"
    calls = []
    monkeypatch.setattr(pipeline, "print_pdf", lambda path, printer=None: calls.append((path, printer)))
    job = ComposeJob(
        pages=_pages(1),
        mode=LayoutMode.ONE_UP,
        out_path=str(out),
        do_print=True,
        copies=3,
        printer="Fake Printer",
    )
    run_compose_job(job)
    assert calls == [(str(out), "Fake Printer")] * 3


def test_run_compose_job_cancel_skips_print(tmp_path, monkeypatch):
    out = tmp_path / "merged_1up.pdf"
    calls = []
    monkeypatch.setattr(pipeline, "print_pdf", lambda path, printer=None: calls.append(path))
    job = ComposeJob(
        pages=_pages(1),
        mode=LayoutMode.ONE_UP,
        out_path=str(out),
        do_print=True,
        copies=2,
    )
    run_compose_job(job, should_cancel=lambda: True)
    assert calls == []


def test_run_compose_job_print_failure_does_not_abort(tmp_path, monkeypatch):
    out = tmp_path / "merged_1up.pdf"
    def _boom(path, printer=None):
        raise RuntimeError("printer on fire")
    monkeypatch.setattr(pipeline, "print_pdf", _boom)
    job = ComposeJob(
        pages=_pages(1),
        mode=LayoutMode.ONE_UP,
        out_path=str(out),
        do_print=True,
        copies=1,
    )
    # 产物已落盘，打印失败只记日志不抛出
    run_compose_job(job)
    assert out.exists()


# ---------------- load_documents ----------------

def test_load_documents_reads_files_and_reports_progress(tmp_path):
    files = []
    for i in range(2):
        f = tmp_path / f"inv{i}.pdf"
        f.write_bytes(_single_page_pdf_bytes())
        files.append(str(f))
    progress = []

    readers, any_ticket, suggested = load_documents(files, progress=lambda i, t, n: progress.append((i, t, n)))

    assert [src for src, _r in readers] == files
    assert all(len(r.pages) == 1 for _src, r in readers)
    assert any_ticket is False
    assert suggested == ""
    assert progress == [(1, 2, "inv0.pdf"), (2, 2, "inv1.pdf")]


def test_load_documents_aggregates_ticket_hints(tmp_path, monkeypatch):
    f = tmp_path / "ticket.pdf"
    f.write_bytes(_single_page_pdf_bytes())
    monkeypatch.setattr(pipeline, "detect_ticket_document", lambda r, max_pages=3: (True, "landscape"))

    readers, any_ticket, suggested = load_documents([str(f)])

    assert len(readers) == 1
    assert any_ticket is True
    assert suggested == "landscape"


# ---------------- process（CLI 批处理） ----------------

def test_process_batch_two_up(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    for i in range(2):
        (src / f"inv{i}.pdf").write_bytes(_single_page_pdf_bytes())
    out = tmp_path / "out"
    out.mkdir()

    pipeline.process(str(src), str(out), do_print=False)

    outputs = sorted(out.glob("*_2up.pdf"))
    assert [f.name for f in outputs] == ["inv0_2up.pdf", "inv1_2up.pdf"]
    for f in outputs:
        # 单页文件 2-up 后仍是 1 页
        assert len(PdfReader(str(f)).pages) == 1


def test_process_output_defaults_to_source_dir(tmp_path):
    f = tmp_path / "one.pdf"
    f.write_bytes(_single_page_pdf_bytes())

    pipeline.process(str(f), None, do_print=False)

    out = tmp_path / "one_2up.pdf"
    assert out.exists()
    assert len(PdfReader(str(out)).pages) == 1
