from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QMessageBox, QVBoxLayout

from database import SessionLocal
from security import authenticate


class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.admin = None
        self.setWindowTitle("Management Console - Sign In")
        self.setMinimumWidth(360)

        title = QLabel("MANAGEMENT CONSOLE")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: 600; padding: 8px;")

        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("Username")
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Password")
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.returnPressed.connect(self.authenticate)

        form = QFormLayout()
        form.addRow("Username", self.username_edit)
        form.addRow("Password", self.password_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Login")
        buttons.accepted.connect(self.authenticate)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(title)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def authenticate(self):
        session = SessionLocal()
        try:
            self.admin = authenticate(
                session,
                self.username_edit.text(),
                self.password_edit.text(),
            )
        finally:
            session.close()
        if self.admin is None:
            QMessageBox.warning(self, "Sign In Failed", "Invalid username or password.")
            self.password_edit.clear()
            self.password_edit.setFocus()
            return
        self.accept()