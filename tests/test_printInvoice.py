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


def test_list_system_printers_returns_names():
    printers = pi.list_system_printers()
    assert isinstance(printers, list)
    assert printers == sorted(set(printers))
    assert all(isinstance(p, str) and p for p in printers)


# ---------------- 可注入策略链 ----------------

def _strategy(name: str, calls: list, result: bool):
    def run() -> bool:
        calls.append(name)
        return result
    return run


def _writable_pdf(tmp_path) -> str:
    pdf = tmp_path / "doc.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    return str(pdf)


def test_print_pdf_uses_first_successful_strategy(monkeypatch, tmp_path):
    monkeypatch.setattr(pi.platform, "system", lambda: "Windows")
    calls: list = []
    strategies = [
        _strategy("a", calls, False),
        _strategy("b", calls, True),
        _strategy("c", calls, True),
    ]
    pi.print_pdf(_writable_pdf(tmp_path), strategies=strategies)
    # 成功即停：失败策略后的第一个成功策略被调用，其余跳过
    assert calls == ["a", "b"]


def test_print_pdf_all_strategies_fail_raises(monkeypatch, tmp_path):
    monkeypatch.setattr(pi.platform, "system", lambda: "Windows")
    calls: list = []
    strategies = [
        _strategy("a", calls, False),
        _strategy("b", calls, False),
    ]
    with pytest.raises(RuntimeError):
        pi.print_pdf(_writable_pdf(tmp_path), strategies=strategies)
    assert calls == ["a", "b"]


def test_default_strategies_direct_print_first_when_printer_given(monkeypatch, tmp_path):
    monkeypatch.setattr(pi, "_find_sumatra", lambda: "C:\\fake\\SumatraPDF.exe")
    monkeypatch.setattr(pi, "_find_edge", lambda: None)
    seen: list = []
    monkeypatch.setattr(pi, "_print_direct", lambda exe, path, printer: seen.append("direct") or True)
    monkeypatch.setattr(pi, "_print_with_sumatra", lambda exe, path: seen.append("dialog") or True)
    monkeypatch.setattr(pi, "_shell_execute_print", lambda path: seen.append("shell") or True)
    monkeypatch.setattr(pi, "_powershell_print", lambda path: seen.append("ps") or True)
    monkeypatch.setattr(pi, "_startfile_print", lambda path: seen.append("startfile") or True)
    monkeypatch.setattr(pi, "_open_with_edge", lambda exe, path: seen.append("edge") or True)
    monkeypatch.setattr(pi, "_open_viewer", lambda path: seen.append("viewer") or True)

    pi.print_pdf(_writable_pdf(tmp_path), printer="My Printer")

    # 指定打印机且 SumatraPDF 可用：直印优先且一次成功
    assert seen == ["direct"]


def test_default_strategies_without_sumatra_falls_back_to_dialog(monkeypatch, tmp_path):
    monkeypatch.setattr(pi, "_find_sumatra", lambda: None)
    monkeypatch.setattr(pi, "_find_edge", lambda: None)
    seen: list = []
    monkeypatch.setattr(pi, "_print_direct", lambda exe, path, printer: seen.append("direct") or True)
    monkeypatch.setattr(pi, "_print_with_sumatra", lambda exe, path: seen.append("dialog") or True)
    monkeypatch.setattr(pi, "_shell_execute_print", lambda path: seen.append("shell") or True)
    monkeypatch.setattr(pi, "_powershell_print", lambda path: seen.append("ps") or True)
    monkeypatch.setattr(pi, "_startfile_print", lambda path: seen.append("startfile") or True)
    monkeypatch.setattr(pi, "_open_with_edge", lambda exe, path: seen.append("edge") or True)
    monkeypatch.setattr(pi, "_open_viewer", lambda path: seen.append("viewer") or True)

    pi.print_pdf(_writable_pdf(tmp_path), printer="My Printer")

    # 未安装 SumatraPDF：跳过直印，落入系统 print 动词
    assert seen == ["shell"]
