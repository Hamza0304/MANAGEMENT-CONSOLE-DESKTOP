import unittest

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from database import Base
from models import Department, PhoneStatus
from services.phone_service import (
    create_phone,
    delete_phone,
    get_phone,
    get_phone_stats,
    get_phones,
    update_phone,
)


class PhoneServiceTests(unittest.TestCase):

    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        Base.metadata.create_all(bind=self.engine)
        self.Session = sessionmaker(bind=self.engine, autoflush=False, autocommit=False, expire_on_commit=False)
        self.session = self.Session()

    def tearDown(self):
        self.session.close()
        Base.metadata.drop_all(bind=self.engine)

    def _make_department(self, name="IT"):
        dept = Department(name=name)
        self.session.add(dept)
        self.session.commit()
        return dept

    def test_create_phone_and_retrieve(self):
        dept = self._make_department()

        phone = create_phone(
            self.session,
            {
                "asset_tag": "P-100",
                "phone_number": "+15551234567",
                "brand": "Apple",
                "model": "iPhone 15",
                "department_id": dept.id,
                "assigned_user": "Alice",
                "location": "HQ",
                "status": PhoneStatus.ASSIGNED.value,
                "operating_system": "iOS",
                "os_version": "17.5",
            },
        )

        self.assertEqual(phone["asset_tag"], "P-100")
        self.assertEqual(phone["department_name"], "IT")

        loaded = get_phone(self.session, phone["id"])
        self.assertEqual(loaded["phone_number"], "+15551234567")

    def test_duplicate_asset_tag_is_rejected(self):
        self._make_department()
        create_phone(
            self.session,
            {
                "asset_tag": "P-100",
                "phone_number": "+15551234567",
                "brand": "Apple",
                "model": "iPhone 15",
            },
        )

        with self.assertRaises(ValueError):
            create_phone(
                self.session,
                {
                    "asset_tag": "P-100",
                    "phone_number": "+15551234568",
                    "brand": "Samsung",
                    "model": "Galaxy S24",
                },
            )

    def test_duplicate_imei_or_serial_is_rejected(self):
        self._make_department()
        create_phone(
            self.session,
            {
                "asset_tag": "P-100",
                "phone_number": "+15551234567",
                "brand": "Apple",
                "model": "iPhone 15",
                "imei_1": "111111111111111",
                "serial_number": "SN-100",
            },
        )

        with self.assertRaises(ValueError):
            create_phone(
                self.session,
                {
                    "asset_tag": "P-101",
                    "phone_number": "+15551234568",
                    "brand": "Samsung",
                    "model": "Galaxy S24",
                    "imei_1": "111111111111111",
                },
            )

    def test_update_phone_reassigns_department_and_status(self):
        dept = self._make_department()
        phone = create_phone(
            self.session,
            {
                "asset_tag": "P-100",
                "phone_number": "+15551234567",
                "brand": "Apple",
                "model": "iPhone 15",
                "department_id": dept.id,
                "status": PhoneStatus.IN_STOCK.value,
            },
        )

        updated = update_phone(
            self.session,
            phone["id"],
            {
                "status": PhoneStatus.ASSIGNED.value,
                "assigned_user": "Bob",
                "location": "Office 1",
            },
        )

        self.assertEqual(updated["status"], PhoneStatus.ASSIGNED.value)
        self.assertEqual(updated["assigned_user"], "Bob")

    def test_get_phone_stats_counts_department_totals(self):
        it = self._make_department("IT")
        finance = self._make_department("Finance")

        create_phone(self.session, {"asset_tag": "P-100", "phone_number": "+15551234567", "brand": "Apple", "model": "iPhone 15", "department_id": it.id, "status": PhoneStatus.ASSIGNED.value})
        create_phone(self.session, {"asset_tag": "P-101", "phone_number": "+15551234568", "brand": "Samsung", "model": "Galaxy S24", "department_id": finance.id, "status": PhoneStatus.IN_STOCK.value})

        stats = get_phone_stats(self.session)

        self.assertEqual(stats["total_phones"], 2)
        self.assertEqual(stats["assigned"], 1)
        self.assertEqual(stats["in_stock"], 1)
        self.assertEqual(stats["department_totals"][0]["department"], "Finance")

    def test_delete_phone_removes_record(self):
        dept = self._make_department()
        phone = create_phone(
            self.session,
            {
                "asset_tag": "P-100",
                "phone_number": "+15551234567",
                "brand": "Apple",
                "model": "iPhone 15",
                "department_id": dept.id,
            },
        )

        delete_phone(self.session, phone["id"])
        self.assertEqual(len(get_phones(self.session)), 0)


if __name__ == "__main__":
    unittest.main()
