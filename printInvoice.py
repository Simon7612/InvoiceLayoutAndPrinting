import os
import platform
import subprocess
import ctypes
import sys

def _sumatra_print_dialog(path: str) -> bool:
    candidates = [
        r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"     
    ]
    # portable next to executable or script
    try:
        base = os.path.dirname(sys.executable)
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(base, "msedge.exe"))
    for exe in candidates:
        if os.path.exists(exe):
            try:
                subprocess.run(
                    [exe, "-print-dialog", "-exit-on-print", path],
                    check=True,
                )
                return True
            except Exception:
                return False
    return False

def _open_viewer(path: str) -> bool:
    try:
        r = ctypes.windll.shell32.ShellExecuteW(None, "open", path, None, None, 1)
        return r > 32
    except Exception:
        pass
    try:
        subprocess.run([
            "powershell",
            "-NoProfile",
            "-Command",
            f"Start-Process -FilePath '{path}'"
        ], check=True)
        return True
    except Exception:
        return False

def _shell_execute_print(path: str) -> bool:
    try:
        r = ctypes.windll.shell32.ShellExecuteW(None, "print", path, None, None, 0)
        return r > 32
    except Exception:
        return False

def _powershell_print(path: str) -> bool:
    try:
        subprocess.run([
            "powershell",
            "-NoProfile",
            "-Command",
            f"Start-Process -FilePath '{path}' -Verb Print"
        ], check=True)
        return True
    except Exception:
        return False

def print_pdf(path: str) -> None:
    if platform.system() != "Windows":
        raise RuntimeError("printing only supported on Windows")
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if _sumatra_print_dialog(path):
        return
    if _powershell_print(path):
        return
    if _shell_execute_print(path):
        return
    try:
        os.startfile(path, "print")
        return
    except Exception:
        pass
    if _open_viewer(path):
        return
    raise RuntimeError("failed to show print dialog; viewer open also failed")