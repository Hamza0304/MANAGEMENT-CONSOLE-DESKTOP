from datetime import datetime
from enum import Enum

from sqlalchemy import Column, Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class PhoneStatus(Enum):
    IN_STOCK = "IN_STOCK"
    ASSIGNED = "ASSIGNED"
    RESERVED = "RESERVED"
    MAINTENANCE = "MAINTENANCE"
    LOST = "LOST"
    RETIRED = "RETIRED"


def normalize_phone_status(value):
    if isinstance(value, PhoneStatus):
        return value.value

    normalized = str(value or PhoneStatus.IN_STOCK.value).strip().upper().replace(" ", "_")

    try:
        return PhoneStatus(normalized).value
    except ValueError as exc:
        raise ValueError(f"Invalid phone status: {value}") from exc


class Phone(Base):
    __tablename__ = "phones"

    id = Column(Integer, primary_key=True)

    asset_tag = Column(String(150), nullable=False, unique=True)
    phone_number = Column(String(50), nullable=False, unique=True)
    brand = Column(String(100), nullable=True)
    model = Column(String(150), nullable=True)

    imei_1 = Column(String(50), nullable=True, unique=True)
    imei_2 = Column(String(50), nullable=True, unique=True)
    serial_number = Column(String(150), nullable=True, unique=True)

    operating_system = Column(String(100), nullable=True)
    os_version = Column(String(100), nullable=True)
    mac_address = Column(String(50), nullable=True)
    ip_address = Column(String(50), nullable=True)

    department_id = Column(Integer, ForeignKey("departments.id"), nullable=True)
    assigned_user = Column(String(150), nullable=True)
    location = Column(String(150), nullable=True)

    status = Column(
        String(30),
        nullable=False,
        default=PhoneStatus.IN_STOCK.value,
    )

    purchase_date = Column(Date, nullable=True)
    warranty_expiry = Column(Date, nullable=True)
    notes = Column(Text, nullable=True)

    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    department = relationship("Department", back_populates="phones")
    audit_logs = relationship("AuditLog", back_populates="phone")

    def __repr__(self):
        return f"<Phone {self.asset_tag}>"
