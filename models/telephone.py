from datetime import datetime
from enum import Enum

from sqlalchemy import Column, Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class TelephoneStatus(Enum):
    IN_STOCK = "IN_STOCK"
    ASSIGNED = "ASSIGNED"
    RESERVED = "RESERVED"
    MAINTENANCE = "MAINTENANCE"
    LOST = "LOST"
    RETIRED = "RETIRED"


def normalize_telephone_status(value):
    if isinstance(value, TelephoneStatus):
        return value.value

    normalized = str(value or TelephoneStatus.IN_STOCK.value).strip().upper().replace(" ", "_")
    try:
        return TelephoneStatus(normalized).value
    except ValueError as exc:
        raise ValueError(f"Invalid telephone status: {value}") from exc


class Telephone(Base):
    __tablename__ = "telephones"

    id = Column(Integer, primary_key=True)
    asset_tag = Column(String(150), nullable=False, unique=True)
    telephone_number = Column(String(50), nullable=False, unique=True)
    brand = Column(String(100), nullable=True)
    model = Column(String(150), nullable=True)
    serial_number = Column(String(150), nullable=True, unique=True)
    mac_address = Column(String(50), nullable=True)
    ip_address = Column(String(50), nullable=True)
    department_id = Column(Integer, ForeignKey("departments.id"), nullable=True)
    assigned_user = Column(String(150), nullable=True)
    location = Column(String(150), nullable=True)
    status = Column(String(30), nullable=False, default=TelephoneStatus.IN_STOCK.value)
    purchase_date = Column(Date, nullable=True)
    notes = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    department = relationship("Department", back_populates="telephones")

    def __repr__(self):
        return f"<Telephone {self.asset_tag}>"
