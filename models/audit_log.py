from datetime import datetime

from sqlalchemy import Column, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True)
    computer_id = Column(Integer, ForeignKey("computers.id"), nullable=True)
    phone_id = Column(Integer, ForeignKey("phones.id"), nullable=True)
    action = Column(String(150), nullable=False)
    previous_status = Column(String(30), nullable=True)
    new_status = Column(String(30), nullable=True)
    department_id = Column(Integer, ForeignKey("departments.id"), nullable=True)
    user = Column(String(150), nullable=True)
    admin = Column(String(150), nullable=True)
    description = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    computer = relationship("Computer", back_populates="audit_logs")
    phone = relationship("Phone", back_populates="audit_logs")
    department = relationship("Department", back_populates="audit_logs")

    def __repr__(self):
        return f"<AuditLog {self.action}>"
