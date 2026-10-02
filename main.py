import logging
import os
import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMessageBox

from database import initialize_database, session_scope, test_connection
from security import ensure_default_admin
from ui.login_dialog import LoginDialog
from ui.main_window import MainWindow

from database import DATABASE_PATH, DATABASE_URL

print("=" * 60)
print("MANAGEMENT CONSOLE DATABASE")
print("=" * 60)
print("DATABASE_PATH:", DATABASE_PATH)
print("DATABASE_URL :", DATABASE_URL)
print("=" * 60)


def main():
    log_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "MANAGEMENT CONSOLE"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / "management_console.log"

    logging.basicConfig(
        filename=str(log_file),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    app = QApplication(sys.argv)
    app.setApplicationName("MANAGEMENT CONSOLE")

    # In a PyInstaller build, bundled data files live under _MEIPASS
    # (the _internal directory in this onedir build). In source runs,
    # use the project directory next to this file.
    resource_root = Path(
        getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    )
    icon_path = resource_root / "icons" / "ManagementConsole.ico"
    if not icon_path.is_file():
        icon_path = Path(__file__).resolve().parent / "icons" / "ManagementConsole.ico"
    if icon_path.is_file():
        app.setWindowIcon(QIcon(str(icon_path)))
    else:
        logging.warning("Application icon file was not found: %s", icon_path)

    try:
        initialize_database()
        test_connection()
        with session_scope() as session:
            ensure_default_admin(session)
    except Exception:
        logging.exception("Unable to initialize the application")
        QMessageBox.critical(None, "Database Connection Failed", "Unable to initialize the configured database.")
        return 1

    login = LoginDialog()
    if login.exec() != LoginDialog.Accepted:
        return 0
    window = MainWindow(login.admin)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
