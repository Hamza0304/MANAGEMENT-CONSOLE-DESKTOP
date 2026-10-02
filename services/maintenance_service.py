from datetime import datetime

from models import Computer, ComputerStatus, Maintenance


def get_maintenance_records(session):
    records = session.query(Maintenance).order_by(Maintenance.recorded_at.desc()).all()
    return [
        {
            "id": record.id,
            "computer_id": record.computer_id,
            "computer_name": record.computer.computer_name,
            "problem": record.problem,
            "solution": record.solution,
            "description": record.description,
            "cost": record.cost,
            "status": record.status,
            "recorded_at": record.recorded_at.isoformat() if record.recorded_at else None,
            "completed_at": record.completed_at.isoformat() if record.completed_at else None,
        }
        for record in records
    ]


def create_maintenance_record(session, payload):
    computer = session.get(Computer, int(payload["computer_id"]))
    if computer is None:
        raise LookupError(f"Computer {payload['computer_id']} was not found")
    record = Maintenance(
        computer_id=computer.id,
        description=payload.get("description"),
        problem=payload.get("problem"),
        solution=payload.get("solution"),
        cost=float(payload.get("cost") or 0),
        status="OPEN",
        recorded_at=datetime.utcnow(),
        created_by=payload.get("created_by") or "admin",
    )
    computer.status = ComputerStatus.MAINTENANCE.value
    session.add(record)
    session.flush()
    return {
        "id": record.id,
        "computer_id": record.computer_id,
        "status": record.status,
        "description": record.description,
    }


def complete_maintenance(session, maintenance_id):
    record = session.get(Maintenance, maintenance_id)
    if record is None:
        raise LookupError(f"Maintenance record {maintenance_id} was not found")
    record.status = "COMPLETED"
    record.solution = record.solution or "Completed"
    record.completed_at = datetime.utcnow()
    record.computer.status = ComputerStatus.IN_STOCK.value
    record.computer.department_id = None
    record.computer.assigned_user = None
    record.computer.location = "IT Storage"
    session.flush()
    return {
        "id": record.id,
        "status": record.status,
        "computer_id": record.computer_id,
    }
