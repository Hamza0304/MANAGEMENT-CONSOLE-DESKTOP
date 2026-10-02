import unittest

import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from database import Base
from models import Department
from services.department_service import get_department
from services.equipment_type_service import (
    create_equipment_type,
    ensure_default_equipment_types,
    get_enabled_equipment_types,
    set_equipment_type_enabled,
)
from services.equipment_asset_service import (
    create_equipment_asset,
    get_equipment_assets,
    import_equipment_assets,
)
from services.equipment_report_service import (
    get_enabled_department_summary,
    get_enabled_equipment_summary,
    get_enabled_report_types,
    get_equipment_report_rows,
)
from services.telephone_service import create_telephone, get_telephones


class EquipmentConfigurationTests(unittest.TestCase):

    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        Base.metadata.create_all(bind=self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.session = self.Session()

    def tearDown(self):
        self.session.close()
        Base.metadata.drop_all(bind=self.engine)

    def test_defaults_and_toggle_are_centralized(self):
        ensure_default_equipment_types(self.session)
        self.session.commit()
        self.assertEqual(
            [item.key for item in get_enabled_equipment_types(self.session)],
            ["computers", "phones"],
        )
        set_equipment_type_enabled(self.session, "telephones", True)
        self.session.commit()
        self.assertIn(
            "telephones",
            [item.key for item in get_enabled_equipment_types(self.session)],
        )

    def test_disabling_type_does_not_delete_telephone_data(self):
        ensure_default_equipment_types(self.session)
        department = Department(name="IT")
        self.session.add(department)
        self.session.commit()
        set_equipment_type_enabled(self.session, "telephones", True)
        create_telephone(
            self.session,
            {
                "asset_tag": "TEL-001",
                "telephone_number": "1001",
                "department_id": department.id,
            },
        )
        self.session.commit()
        set_equipment_type_enabled(self.session, "telephones", False)
        self.session.commit()
        self.assertEqual(len(get_telephones(self.session)), 1)

    def test_generic_printer_management_and_excel_import(self):
        ensure_default_equipment_types(self.session)
        create_equipment_asset(
            self.session,
            "printers",
            {"asset_tag": "PRN-001", "brand": "HP", "model": "LaserJet"},
        )
        result = import_equipment_assets(
            self.session,
            "printers",
            pd.DataFrame([{"asset_tag": "PRN-002", "brand": "Canon"}]),
        )
        self.session.commit()
        self.assertEqual(result["created"], 1)
        self.assertEqual(len(get_equipment_assets(self.session, "printers")), 2)

    def test_reports_follow_enabled_types_without_deleting_records(self):
        ensure_default_equipment_types(self.session)
        department = Department(name="Informatique")
        self.session.add(department)
        self.session.commit()
        set_equipment_type_enabled(self.session, "tablets", True)
        create_equipment_asset(
            self.session,
            "tablets",
            {
                "asset_tag": "TAB-001",
                "brand": "Apple",
                "model": "iPad",
                "department_id": department.id,
                "status": "ASSIGNED",
            },
        )
        self.session.commit()
        self.assertIn("tablets", [item.key for item in get_enabled_report_types(self.session)])
        self.assertEqual(len(get_equipment_report_rows(self.session, "tablets")), 1)
        self.assertEqual(get_enabled_department_summary(self.session)[0]["Tablets"], 1)
        set_equipment_type_enabled(self.session, "tablets", False)
        self.session.commit()
        self.assertNotIn("tablets", [item.key for item in get_enabled_report_types(self.session)])
        self.assertEqual(len(get_equipment_assets(self.session, "tablets")), 1)
        set_equipment_type_enabled(self.session, "tablets", True)
        self.session.commit()
        self.assertEqual(get_enabled_equipment_summary(self.session)[-1]["name"], "Tablets")

    def test_new_equipment_type_uses_generic_registry(self):
        equipment_type = create_equipment_type(
            self.session,
            "projectors",
            "Projectors",
            "Meeting room projectors",
            enabled=True,
        )
        self.session.commit()
        self.assertEqual(equipment_type.key, "projectors")
        create_equipment_asset(
            self.session,
            "projectors",
            {"asset_tag": "PROJ-001", "brand": "Epson"},
        )
        self.session.commit()
        self.assertEqual(len(get_equipment_report_rows(self.session, "projectors")), 1)

    def test_department_sections_follow_enabled_registry(self):
        ensure_default_equipment_types(self.session)
        department = Department(name="Informatique")
        self.session.add(department)
        self.session.commit()
        set_equipment_type_enabled(self.session, "tablets", True)
        create_equipment_asset(
            self.session,
            "tablets",
            {"asset_tag": "TAB-001", "department_id": department.id},
        )
        self.session.commit()
        result = get_department(self.session, department.id)
        self.assertEqual(
            [section["key"] for section in result["equipment_sections"]],
            ["tablets"],
        )
        set_equipment_type_enabled(self.session, "tablets", False)
        self.session.commit()
        self.assertEqual(get_department(self.session, department.id)["equipment_sections"], [])
        set_equipment_type_enabled(self.session, "tablets", True)
        self.session.commit()
        self.assertEqual(
            get_department(self.session, department.id)["equipment_sections"][0]["key"],
            "tablets",
        )

    def test_generic_equipment_department_and_status_filters(self):
        ensure_default_equipment_types(self.session)
        first = Department(name="IT")
        second = Department(name="Reception")
        self.session.add_all([first, second])
        self.session.commit()
        create_equipment_asset(
            self.session,
            "printers",
            {"asset_tag": "PRN-IT", "department_id": first.id, "status": "ASSIGNED"},
        )
        create_equipment_asset(
            self.session,
            "printers",
            {"asset_tag": "PRN-REC", "department_id": second.id, "status": "IN_STOCK"},
        )
        self.session.commit()
        self.assertEqual(
            len(get_equipment_assets(self.session, "printers", {"department_id": first.id})),
            1,
        )
        self.assertEqual(
            len(get_equipment_assets(self.session, "printers", {"status": "IN_STOCK"})),
            1,
        )

    def test_telephone_department_and_status_filters(self):
        first = Department(name="IT")
        second = Department(name="Reception")
        self.session.add_all([first, second])
        self.session.commit()
        create_telephone(
            self.session,
            {
                "asset_tag": "TEL-IT",
                "telephone_number": "201",
                "department_id": first.id,
                "status": "ASSIGNED",
            },
        )
        create_telephone(
            self.session,
            {
                "asset_tag": "TEL-REC",
                "telephone_number": "202",
                "department_id": second.id,
                "status": "IN_STOCK",
            },
        )
        self.session.commit()
        self.assertEqual(len(get_telephones(self.session, {"department_id": first.id})), 1)
        self.assertEqual(len(get_telephones(self.session, {"status": "IN_STOCK"})), 1)


if __name__ == "__main__":
    unittest.main()
