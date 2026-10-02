from datetime import datetime

from sqlalchemy import Column, DateTime, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class Department(Base):
    __tablename__ = "departments"

    id = Column(Integer, primary_key=True)
    name = Column(String(150), nullable=False, unique=True)
    description = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    computers = relationship(
        "Computer",
        back_populates="department",
    )
    phones = relationship(
        "Phone",
        back_populates="department",
    )
    telephones = relationship(
        "Telephone",
        back_populates="department",
    )
    equipment_assets = relationship(
        "EquipmentAsset",
        back_populates="department",
    )
    audit_logs = relationship("AuditLog", back_populates="department")

    @property
    def computer_count(self):
        return len(self.computers)

    def __repr__(self):
        return f"<Department {self.name}>"
