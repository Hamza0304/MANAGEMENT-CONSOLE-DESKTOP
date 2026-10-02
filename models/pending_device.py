from datetime import datetime

from sqlalchemy import Column, DateTime, Integer, String, Text

from database import Base


class PendingDevice(Base):
    __tablename__ = "pending_devices"

    id = Column(Integer, primary_key=True)
    computer_name = Column(String(100), nullable=False, unique=True)
    manufacturer = Column(String(100), nullable=True)
    model = Column(String(100), nullable=True)
    serial_number = Column(String(150), nullable=True)
    operating_system = Column(String(100), nullable=True)
    os_version = Column(String(100), nullable=True)
    cpu = Column(String(150), nullable=True)
    ram = Column(String(50), nullable=True)
    storage = Column(String(100), nullable=True)
    ip_address = Column(String(50), nullable=True)
    mac_address = Column(String(50), nullable=True)
    installed_software = Column(Text, nullable=True)
    last_seen = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)