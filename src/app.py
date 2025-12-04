from PySide6.QtWidgets import QApplication
import sys
from pathlib import Path

if __package__:
    from .ui.main_window import MainWindow
else:
    sys.path.append(str(Path(__file__).resolve().parent))
    from ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
