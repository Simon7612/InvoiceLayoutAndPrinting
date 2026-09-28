import ctypes
import logging
import os
import platform
import subprocess
import sys

logger = logging.getLogger(__name__)


def _exe_dir() -> str:
    """打包后 exe 所在目录（便携程序可放在旁边）。"""
    try:
        return os.path.dirname(sys.executable)
    except Exception:
        return os.path.dirname(os.path.abspath(__file__))


def _query_app_path(exe_name: str):
    """从注册表 App Paths 查询可执行文件路径（Windows 官方推荐做法）。"""
    try:
        import winreg

        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            value = winreg.QueryValue(key, None)
            if value and os.path.exists(value):
                return value
    except Exception:
        pass
    return None


def _find_sumatra():
    """SumatraPDF 支持 -print-dialog 参数，是唯一真正的命令行打印对话框来源。"""
    for base in (
        _exe_dir(),
        os.path.dirname(os.path.abspath(__file__)),
        r"C:\Program Files\SumatraPDF",
        r"C:\Program Files (x86)\SumatraPDF",
    ):
        exe = os.path.join(base, "SumatraPDF.exe")
        if os.path.exists(exe):
            return exe
    return _query_app_path("SumatraPDF.exe")


def _find_edge():
    for exe in (
        _query_app_path("msedge.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.join(_exe_dir(), "msedge.exe"),
    ):
        if exe and os.path.exists(exe):
            return exe
    return None


def _spawn(args: list) -> bool:
    """非阻塞启动 GUI 程序；阻塞等待会让界面/线程卡在打印窗口上。"""
    try:
        subprocess.Popen(args)
        return True
    except Exception as e:
        logger.warning("启动 %s 失败: %s", args[0] if args else "?", e)
        return False


def _print_with_sumatra(exe: str, path: str) -> bool:
    return _spawn([exe, "-print-dialog", "-exit-on-print", os.path.abspath(path)])


def _open_with_edge(exe: str, path: str) -> bool:
    # Edge 没有打印 CLI 参数；只能打开其查看器由用户在窗口内打印
    return _spawn([exe, os.path.abspath(path)])


def _shell_execute_print(path: str) -> bool:
    """调用系统 "print" 动词（默认 PDF 程序自己的打印对话框）。"""
    try:
        r = ctypes.windll.shell32.ShellExecuteW(None, "print", path, None, None, 0)
        return r > 32
    except Exception as e:
        logger.debug("ShellExecuteW print 失败: %s", e)
        return False


def _powershell_print(path: str) -> bool:
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"Start-Process -FilePath '{path}' -Verb Print",
            ],
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return True
    except Exception as e:
        logger.debug("PowerShell print 失败: %s", e)
        return False


def _open_viewer(path: str) -> bool:
    try:
        r = ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 1)
        return r > 32
    except Exception as e:
        logger.debug("ShellExecuteW open 失败: %s", e)
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"Start-Process -FilePath '{path}'",
            ],
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return True
    except Exception as e:
        logger.debug("PowerShell Start-Process 失败: %s", e)
        return False


def print_pdf(path: str) -> None:
    """弹出打印对话框（或兜底打开查看器），不阻塞调用方。"""
    if platform.system() != "Windows":
        raise RuntimeError("printing only supported on Windows")
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    # 1) SumatraPDF：真正的命令行打印对话框
    sumatra = _find_sumatra()
    if sumatra and _print_with_sumatra(sumatra, path):
        return
    # 2) 系统 "print" 动词：默认 PDF 程序的打印对话框
    if _shell_execute_print(path):
        return
    if _powershell_print(path):
        return
    try:
        os.startfile(path, "print")
        return
    except Exception as e:
        logger.debug("os.startfile print 失败: %s", e)
    # 3) 兜底：打开查看器由用户手动打印
    edge = _find_edge()
    if edge and _open_with_edge(edge, path):
        return
    if _open_viewer(path):
        return
    raise RuntimeError("failed to show print dialog; viewer open also failed")
