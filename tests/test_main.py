"""main.py 参数分发测试（批处理编排已在 pipeline，见 test_pipeline.py）。"""
import sys

import main as main_mod
import pipeline


def test_main_cli_dispatches_to_process(monkeypatch):
    calls = {}
    monkeypatch.setattr(
        pipeline, "process",
        lambda i, o, p: calls.update(input=i, output=o, do_print=p),
    )
    monkeypatch.setattr(sys, "argv", ["main.py", "-i", "a", "-o", "b", "--no-print"])
    main_mod.main()
    assert calls == {"input": "a", "output": "b", "do_print": False}


def test_main_print_enabled_by_default(monkeypatch):
    calls = {}
    monkeypatch.setattr(
        pipeline, "process",
        lambda i, o, p: calls.update(do_print=p),
    )
    monkeypatch.setattr(sys, "argv", ["main.py", "-i", "a"])
    main_mod.main()
    assert calls == {"do_print": True}


def test_main_gui_flag_and_missing_input_enter_gui(monkeypatch):
    calls = {"gui": 0, "process": 0}
    monkeypatch.setattr(main_mod, "run_gui", lambda: calls.__setitem__("gui", calls["gui"] + 1))
    monkeypatch.setattr(
        pipeline, "process",
        lambda *a: calls.__setitem__("process", calls["process"] + 1),
    )
    monkeypatch.setattr(sys, "argv", ["main.py", "--gui"])
    main_mod.main()
    # 无 -i 时也直接进 GUI
    monkeypatch.setattr(sys, "argv", ["main.py"])
    main_mod.main()
    assert calls == {"gui": 2, "process": 0}
