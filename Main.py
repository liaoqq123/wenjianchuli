import importlib.util
import os
import sys
from pathlib import Path
import tkinter as tk


def ensure_project_environment():
    if importlib.util.find_spec("ttkbootstrap") is not None:
        return

    project_dir = Path(__file__).resolve().parent
    for python_path in (
        project_dir / ".venv" / "Scripts" / "python.exe",
        project_dir / "venv" / "Scripts" / "python.exe",
    ):
        if not python_path.exists():
            continue

        if python_path.resolve() == Path(sys.executable).resolve():
            continue

        os.execv(str(python_path), [str(python_path), *sys.argv])

    raise ModuleNotFoundError(
        "没有找到 ttkbootstrap。请使用 run.bat 启动，或先运行：python -m pip install -r requirements.txt"
    )


ensure_project_environment()

import ttkbootstrap as ttk

from Process import ProcessFrame


class MainPage:
    def __init__(self, master: tk.Tk):
        self.window = master
        self.window.title("文件处理器 1.0.0")
        self.window.geometry("960x540")
        self.root()

    def root(self):
        self.process = ProcessFrame(self.window)
        self.process.pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    window = ttk.Window(themename="flatly")
    MainPage(window)
    window.mainloop()
