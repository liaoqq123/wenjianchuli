import importlib.util
import os
import sys
from pathlib import Path


def ensure_project_environment():
    if getattr(sys, "frozen", False):
        return

    # 双击启动时优先切到项目自带虚拟环境，避免系统 Python 缺少 PySide6。
    if importlib.util.find_spec("PySide6") is not None:
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
        "没有找到 PySide6。请使用 run.bat 启动，或先运行：python -m pip install -r requirements.txt"
    )


ensure_project_environment()

# 环境确认后再导入界面库，确保缺依赖时能先重启到虚拟环境。
from PySide6.QtWidgets import QApplication, QMainWindow

from Process import ProcessFrame


class MainPage(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文件处理器 1.0.0")
        self.resize(960, 540)
        self.setCentralWidget(ProcessFrame(self))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainPage()
    window.show()
    sys.exit(app.exec())
