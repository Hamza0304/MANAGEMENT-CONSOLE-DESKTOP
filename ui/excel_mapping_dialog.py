from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QInputDialog,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from database import SessionLocal
from models import Computer, Department
from services.computer_service import create_computer
from services.department_service import get_departments
from services.excel_mapping_service import (
    TARGET_FIELD_LABELS,
    TARGET_FIELDS,
    build_preview_rows,
    detect_auto_mapping,
    load_mapping_profiles,
    save_mapping_profile,
    build_mapping_from_profile,
    detect_mapping_profile,
    normalize_column_name,
)


class ExcelMappingDialog(QDialog):
    def __init__(self, parent, dataframe, detected_mappings=None, profile=None):
        super().__init__(parent)
        self.dataframe = dataframe
        self.detected_mappings = self._normalize_detected_mappings(
            detected_mappings or detect_auto_mapping(list(dataframe.columns)),
            list(dataframe.columns),
        )
        self.profile = profile
        self.profile_name = profile.get("name") if profile else None
        self.rows = []
        self.preview_rows = []
        self.summary = None
        self.source_columns = list(dataframe.columns)
        self.departments = []
        self.department_lookup = {}
        self.existing_computer_names = set()
        self.existing_serial_numbers = set()

        self.setWindowTitle("Excel Mapping Wizard")
        self.resize(1200, 820)

        self.top_label = QLabel(
            "Map Excel columns to Management Console fields."
        )
        self.top_label.setWordWrap(True)

        self.status_label = QLabel()
        self.status_label.setWordWrap(True)

        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(3)
        self.mapping_table.itemChanged.connect(self._handle_mapping_item_changed)
        self.mapping_table.setHorizontalHeaderLabels([
            "Excel Column",
            "Management Console Field",
            "Status",
        ])
        self.mapping_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapping_table.setAlternatingRowColors(True)

        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.preview_table.setMinimumHeight(260)

        self.summary_label = QLabel()
        self.summary_label.setWordWrap(True)

        button_row = QHBoxLayout()

        self.add_column_button = QPushButton("Add Column")
        self.add_column_button.clicked.connect(self.add_column)

        self.preview_button = QPushButton("Preview")
        self.preview_button.clicked.connect(self.preview_import)

        self.save_profile_button = QPushButton("Save Profile")
        self.save_profile_button.clicked.connect(self.save_profile)

        self.import_button = QPushButton("Import Valid Rows")
        self.import_button.clicked.connect(self.import_rows)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)

        button_row.addWidget(self.add_column_button)
        button_row.addWidget(self.save_profile_button)
        button_row.addStretch()
        button_row.addWidget(self.preview_button)
        button_row.addWidget(self.import_button)
        button_row.addWidget(self.cancel_button)

        self.main_layout = QVBoxLayout(self)
        self.main_layout.addWidget(self.top_label)
        self.main_layout.addWidget(self.status_label)
        self.main_layout.addWidget(self.mapping_table)
        self.main_layout.addWidget(self.preview_table)
        self.main_layout.addWidget(self.summary_label)
        self.main_layout.addLayout(button_row)

        self._load_existing_context()
        self._build_initial_rows()
        self._refresh_mapping_table()
        self._update_status_label()
        self._populate_preview_table([])

    def _load_existing_context(self):
        session = SessionLocal()
        try:
            departments = get_departments(session)
            self.department_lookup = {
                normalize_column_name(item["name"]): item
                for item in departments
                if item.get("name")
            }

            computers = session.query(Computer).all()
            self.existing_computer_names = {
                normalize_column_name(computer.computer_name)
                for computer in computers
                if computer.computer_name
            }
            self.existing_serial_numbers = {
                normalize_column_name(computer.serial_number)
                for computer in computers
                if computer.serial_number
            }
        finally:
            session.close()

    def _normalize_detected_mappings(self, detected_mappings, columns):
        if not detected_mappings:
            return {}

        if any(key in columns for key in detected_mappings):
            return detected_mappings

        normalized = {}
        for target_field, source_column in detected_mappings.items():
            normalized[source_column] = target_field
        return normalized

    def _build_initial_rows(self):
        self.rows = []
        for column in self.source_columns:
            self.rows.append(
                {
                    "source_column": column,
                    "target_field": self.detected_mappings.get(column),
                    "added": False,
                }
            )

    def _refresh_mapping_table(self):
        self.mapping_table.blockSignals(True)
        try:
            self.mapping_table.setRowCount(len(self.rows))
            self.mapping_table.setColumnWidth(0, 260)
            self.mapping_table.setColumnWidth(1, 280)
            self.mapping_table.setColumnWidth(2, 120)

            for row_index, row in enumerate(self.rows):
                source_text = row.get("source_column") or "Added Column"
                source_item = QTableWidgetItem(source_text)
                if row.get("added"):
                    source_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled)
                else:
                    source_item.setFlags(source_item.flags() & ~Qt.ItemIsEditable)
                self.mapping_table.setItem(row_index, 0, source_item)

                target_combo = QComboBox()
                target_combo.blockSignals(True)
                try:
                    target_combo.addItem("Not Mapped", "")
                    for field in TARGET_FIELDS:
                        target_combo.addItem(field["label"], field["key"])

                    selected_target = row.get("target_field")
                    if selected_target:
                        index = target_combo.findData(selected_target)
                        if index >= 0:
                            target_combo.setCurrentIndex(index)
                        else:
                            target_combo.setCurrentIndex(0)
                    else:
                        target_combo.setCurrentIndex(0)
                finally:
                    target_combo.blockSignals(False)

                target_combo.currentIndexChanged.connect(
                    lambda index, row_index=row_index, combo=target_combo: self._handle_target_change(row_index, combo, index)
                )
                self.mapping_table.setCellWidget(row_index, 1, target_combo)

                status = self._row_status(row)
                item = QTableWidgetItem(status)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.mapping_table.setItem(row_index, 2, item)
        finally:
            self.mapping_table.blockSignals(False)

    def _row_status(self, row):
        target = row.get("target_field")
        if row.get("added"):
            return "Added"
        if target and self.detected_mappings.get(row.get("source_column")) == target:
            return "Detected"
        if target:
            return "Mapped"
        return "Unmapped"

    def _handle_mapping_item_changed(self, item):
        if item.column() != 0:
            return

        row_index = item.row()
        if row_index >= len(self.rows):
            return

        row = self.rows[row_index]
        if not row.get("added"):
            return

        new_source = item.text().strip()
        row["source_column"] = new_source
        self._update_status_label()

    def _update_row_status(self, row_index):
        if row_index >= len(self.rows):
            return

        row = self.rows[row_index]
        status_item = self.mapping_table.item(row_index, 2)
        if status_item is None:
            status_item = QTableWidgetItem()
            self.mapping_table.setItem(row_index, 2, status_item)

        status_item.setText(self._row_status(row))
        status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)

    def _handle_target_change(self, row_index, combo, index):
        target_value = combo.currentData()
        row = self.rows[row_index]
        selected_targets = []

        for idx, current_row in enumerate(self.rows):
            if idx == row_index:
                continue
            if current_row.get("target_field"):
                selected_targets.append(current_row.get("target_field"))

        if target_value and target_value in selected_targets:
            QMessageBox.warning(
                self,
                "Duplicate Field Mapping",
                "That Management Console field is already mapped to another Excel column. Please choose a different target field.",
            )
            combo.blockSignals(True)
            try:
                combo.setCurrentIndex(0)
            finally:
                combo.blockSignals(False)
            row["target_field"] = None
            self._update_row_status(row_index)
            self._update_status_label()
            return

        row["target_field"] = target_value
        self._update_row_status(row_index)
        self._update_status_label()

    def _update_status_label(self):
        detected = [row for row in self.rows if row.get("target_field")]
        if self.profile_name:
            self.status_label.setText(
                f"Existing mapping profile detected: {self.profile_name}"
            )
        else:
            self.status_label.setText(
                f"Auto-detected mappings: {len(detected)}"
            )

    def add_column(self):
        available_targets = []
        used_targets = {row.get("target_field") for row in self.rows if row.get("target_field")}

        for field in TARGET_FIELDS:
            if field["key"] not in used_targets:
                available_targets.append(field)

        if not available_targets:
            QMessageBox.information(
                self,
                "No More Columns",
                "All available target fields are already mapped.",
            )
            return

        target_labels = [field["label"] for field in available_targets]
        selected_label, ok = QInputDialog.getItem(
            self,
            "Add Column",
            "Select a Management Console field to add:",
            target_labels,
            0,
            False,
        )

        if not ok:
            return

        selected_target = next(
            field["key"]
            for field in available_targets
            if field["label"] == selected_label
        )

        self.rows.append(
            {
                "source_column": "",
                "target_field": selected_target,
                "added": True,
            }
        )

        self._refresh_mapping_table()
        self._update_status_label()

    def save_profile(self):
        current_mapping = {}
        for row in self.rows:
            if not row.get("target_field"):
                continue
            if row.get("source_column"):
                current_mapping[row["target_field"]] = row["source_column"]
            else:
                current_mapping[row["target_field"]] = None

        profile_name, ok = QInputDialog.getText(
            self,
            "Save Mapping Profile",
            "Enter a profile name:",
        )

        if not ok:
            return

        try:
            save_mapping_profile(profile_name, current_mapping)
            QMessageBox.information(
                self,
                "Profile Saved",
                f"Saved mapping profile: {profile_name}",
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Profile Save Failed",
                f"Could not save the profile.\n\n{exc}",
            )

    def preview_import(self):
        current_mapping = {}
        for row in self.rows:
            if row.get("target_field"):
                current_mapping[row["target_field"]] = row.get("source_column") or None

        if "computer_name" not in current_mapping or not current_mapping.get("computer_name"):
            QMessageBox.warning(
                self,
                "Computer Name is required",
                "Computer Name is required before importing.",
            )
            return

        try:
            preview_rows, summary = build_preview_rows(
                self.dataframe,
                current_mapping,
                self.department_lookup,
                self.existing_computer_names,
                self.existing_serial_numbers,
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Preview Error",
                str(exc),
            )
            return

        self.preview_rows = preview_rows
        self.summary = summary
        self._populate_preview_table(preview_rows)

        self.summary_label.setText(
            f"Total Excel Rows: {summary['total_rows']} | "
            f"Valid Rows: {summary['valid_rows']} | "
            f"Warnings: {summary['warning_rows']} | "
            f"Errors: {summary['error_rows']}"
        )

    def _populate_preview_table(self, preview_rows):
        self.preview_table.clear()
        self.preview_table.setRowCount(0)

        if not preview_rows:
            self.preview_table.setColumnCount(1)
            self.preview_table.setHorizontalHeaderLabels(["Preview"])
            self.preview_table.setRowCount(1)
            self.preview_table.setItem(0, 0, QTableWidgetItem("Click Preview to build the import preview."))
            return

        mapped_target_keys = []
        for field in TARGET_FIELDS:
            if field["key"] in self._current_target_map() and field["key"] != "status":
                mapped_target_keys.append(field["key"])

        headers = [TARGET_FIELD_LABELS.get(key, key) for key in mapped_target_keys]
        headers.extend(["Status", "Messages"])
        self.preview_table.setColumnCount(len(headers))
        self.preview_table.setHorizontalHeaderLabels(headers)
        self.preview_table.setRowCount(len(preview_rows))

        for row_index, preview_row in enumerate(preview_rows):
            data = preview_row.get("data", {})
            payload = preview_row.get("payload", {})
            column_index = 0

            for key in mapped_target_keys:
                value = data.get(key)
                if value is None and key in payload:
                    value = payload.get(key)
                if value is None:
                    display = ""
                else:
                    display = str(value)
                self.preview_table.setItem(row_index, column_index, QTableWidgetItem(display))
                column_index += 1

            status_item = QTableWidgetItem(preview_row.get("display_status", ""))
            self.preview_table.setItem(row_index, column_index, status_item)
            column_index += 1

            message_text = "\n".join(preview_row.get("messages", []))
            self.preview_table.setItem(row_index, column_index, QTableWidgetItem(message_text))

        self.preview_table.resizeColumnsToContents()

    def _current_target_map(self):
        current_map = {}
        for row in self.rows:
            if row.get("target_field"):
                current_map[row["target_field"]] = row.get("source_column") or None
        return current_map

    def import_rows(self):
        if not self.preview_rows:
            QMessageBox.warning(
                self,
                "Preview Required",
                "Please generate the preview before importing.",
            )
            return

        rows_to_import = [row for row in self.preview_rows if row.get("status") != "error"]

        if not rows_to_import:
            QMessageBox.warning(
                self,
                "No Rows to Import",
                "All rows contain errors and cannot be imported.",
            )
            return

        confirmation = QMessageBox.question(
            self,
            "Confirm Excel Import",
            (
                f"Total Excel Rows: {self.summary['total_rows']}\n"
                f"Valid Rows: {self.summary['valid_rows']}\n"
                f"Warnings: {self.summary['warning_rows']}\n"
                f"Errors: {self.summary['error_rows']}\n\n"
                "Import the valid and warning rows?"
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if confirmation != QMessageBox.Yes:
            return

        imported_count = 0
        skipped_count = 0
        session = SessionLocal()

        try:
            for preview_row in rows_to_import:
                payload = self._prepare_payload(preview_row.get("payload", {}))

                department_name = payload.pop("department_name", None)
                if department_name:
                    department = session.query(Department).filter(
                        Department.name == department_name
                    ).first()

                    if department is None:
                        department = Department(
                            name=department_name,
                            description="Created during Excel import",
                        )
                        session.add(department)
                        session.flush()

                    payload["department_id"] = department.id

                existing = session.query(Computer).filter(
                    Computer.computer_name == payload.get("computer_name")
                ).first()
                if existing:
                    skipped_count += 1
                    continue

                serial_number = payload.get("serial_number")
                if serial_number:
                    existing_serial = session.query(Computer).filter(
                        Computer.serial_number == serial_number
                    ).first()
                    if existing_serial:
                        skipped_count += 1
                        continue

                create_computer(session, payload)
                imported_count += 1

            session.commit()
        except Exception as exc:
            session.rollback()
            QMessageBox.critical(
                self,
                "Import Error",
                f"The import failed and no changes were saved.\n\n{exc}",
            )
            return
        finally:
            session.close()

        QMessageBox.information(
            self,
            "Import Completed",
            (
                f"Excel import completed successfully.\n\n"
                f"Imported: {imported_count}\n"
                f"Skipped: {skipped_count}"
            ),
        )
        self.accept()

    def _prepare_payload(self, payload):
        prepared = dict(payload)

        if prepared.get("purchase_date"):
            try:
                prepared["purchase_date"] = datetime.strptime(
                    prepared["purchase_date"], "%Y-%m-%d"
                ).date()
            except Exception:
                prepared["purchase_date"] = None

        if prepared.get("location") is None:
            prepared["location"] = "IT Storage"

        return prepared
