"""printInvoice 打印链路的守卫逻辑测试。

不在测试中真正启动外部打印程序，只验证查找与入参守卫。
"""
import os

import pytest

import printInvoice as pi


def test_print_pdf_rejects_non_windows(monkeypatch):
    monkeypatch.setattr(pi.platform, "system", lambda: "Linux")
    with pytest.raises(RuntimeError):
        pi.print_pdf("whatever.pdf")


def test_print_pdf_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        pi.print_pdf(str(tmp_path / "nope.pdf"))


def test_find_edge_returns_existing_path_or_none():
    exe = pi._find_edge()
    assert exe is None or os.path.exists(exe)


def test_find_sumatra_returns_existing_path_or_none():
    exe = pi._find_sumatra()
    assert exe is None or os.path.exists(exe)


def test_query_app_path_returns_existing_path_or_none():
    result = pi._query_app_path("msedge.exe")
    assert result is None or os.path.exists(result)
