from datetime import datetime
from enum import Enum

from sqlalchemy import Column, Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class ComputerStatus(Enum):
    IN_STOCK = "IN_STOCK"
    ASSIGNED = "ASSIGNED"
    RESERVED = "RESERVED"
    MAINTENANCE = "MAINTENANCE"
    RETIRED = "RETIRED"


def normalize_status(value):
    if isinstance(value, ComputerStatus):
        return value.value
    normalized = str(value or ComputerStatus.IN_STOCK.value).strip().upper().replace(" ", "_")
    try:
        return ComputerStatus(normalized).value
    except ValueError as exc:
        raise ValueError(f"Invalid computer status: {value}") from exc


class Computer(Base):
    __tablename__ = "computers"

    id = Column(Integer, primary_key=True)

    # Basic information
    computer_name = Column(String(100), nullable=False, unique=True)
    manufacturer = Column(String(100), nullable=True)
    model = Column(String(100), nullable=True)
    serial_number = Column(String(150), unique=True, nullable=True)

    # Operating system
    operating_system = Column(String(100), nullable=True)
    os_version = Column(String(100), nullable=True)

    # Hardware
    cpu = Column(String(150), nullable=True)
    ram = Column(String(50), nullable=True)
    storage = Column(String(100), nullable=True)

    # Network
    ip_address = Column(String(50), nullable=True)
    mac_address = Column(String(50), nullable=True)

    # Assignment
    department_id = Column(
        Integer,
        ForeignKey("departments.id"),
        nullable=True
    )

    assigned_user = Column(String(150), nullable=True)
    location = Column(String(150), nullable=True)

    # Status
    status = Column(
        String(30),
        nullable=False,
        default=ComputerStatus.IN_STOCK.value,
    )

    # Additional information
    purchase_date = Column(Date, nullable=True)
    warranty = Column(String(100), nullable=True)
    notes = Column(Text, nullable=True)
    installed_software = Column(Text, nullable=True)

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )

    last_seen = Column(DateTime, nullable=True)

    # Relationship
    department = relationship(
        "Department",
        back_populates="computers",
    )
    maintenance_records = relationship("Maintenance", back_populates="computer")
    audit_logs = relationship("AuditLog", back_populates="computer")

    def __repr__(self):
        return f"<Computer {self.computer_name}>"