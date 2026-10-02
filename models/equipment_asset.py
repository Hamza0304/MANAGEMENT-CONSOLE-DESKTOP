from datetime import datetime

from sqlalchemy import Column, Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class EquipmentAsset(Base):
    __tablename__ = "equipment_assets"

    id = Column(Integer, primary_key=True)
    equipment_type_id = Column(Integer, ForeignKey("equipment_types.id"), nullable=False, index=True)
    asset_tag = Column(String(150), nullable=False)
    name = Column(String(150), nullable=True)
    brand = Column(String(100), nullable=True)
    model = Column(String(150), nullable=True)
    serial_number = Column(String(150), nullable=True)
    ip_address = Column(String(50), nullable=True)
    mac_address = Column(String(50), nullable=True)
    department_id = Column(Integer, ForeignKey("departments.id"), nullable=True)
    assigned_user = Column(String(150), nullable=True)
    location = Column(String(150), nullable=True)
    status = Column(String(30), nullable=False, default="IN_STOCK")
    purchase_date = Column(Date, nullable=True)
    notes = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    equipment_type = relationship("EquipmentType")
    department = relationship("Department", back_populates="equipment_assets")

    def __repr__(self):
        return f"<EquipmentAsset {self.asset_tag}>"
