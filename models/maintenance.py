from datetime import datetime

from sqlalchemy import Column, DateTime, Float, ForeignKey, Integer, String, Text
from sqlalchemy.orm import relationship

from database import Base


class Maintenance(Base):
    __tablename__ = "maintenance_records"

    id = Column(Integer, primary_key=True)
    computer_id = Column(Integer, ForeignKey("computers.id"), nullable=False)
    description = Column(Text, nullable=True)
    problem = Column(Text, nullable=True)
    solution = Column(Text, nullable=True)
    cost = Column(Float, default=0.0)
    status = Column(String(30), default="OPEN", nullable=False)
    recorded_at = Column(DateTime, default=datetime.utcnow)
    completed_at = Column(DateTime, nullable=True)
    created_by = Column(String(150), nullable=True)

    computer = relationship("Computer", back_populates="maintenance_records")

    def __repr__(self):
        return f"<Maintenance {self.id}>"
