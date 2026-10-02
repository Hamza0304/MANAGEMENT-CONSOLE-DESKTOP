from datetime import datetime

from sqlalchemy import Boolean, Column, DateTime, Integer, String, Text

from database import Base


DEFAULT_EQUIPMENT_TYPES = (
    ("computers", "Computers", "Desktop PCs and laptops", True),
    ("phones", "Phones", "Company mobile phones", True),
    ("telephones", "Telephones", "Desk phones / office telephones", False),
    ("printers", "Printers", "Office printers", False),
    ("servers", "Servers", "Company servers", False),
    ("network_equipment", "Network Equipment", "Switches, routers, access points", False),
    ("tablets", "Tablets", "Company tablets", False),
    ("monitors", "Monitors", "Company monitors", False),
)


class EquipmentType(Base):
    __tablename__ = "equipment_types"

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)
    key = Column(String(100), nullable=False, unique=True, index=True)
    description = Column(Text, nullable=True)
    enabled = Column(Boolean, nullable=False, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def __repr__(self):
        return f"<EquipmentType {self.key} enabled={self.enabled}>"
