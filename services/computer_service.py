from datetime import datetime

from models import AuditLog, Computer, ComputerStatus, Department, PendingDevice
from models.computer import normalize_status

ONLINE_THRESHOLD_MINUTES = 5


def get_connection_status(last_seen, now=None):
    if last_seen is None:
        return "NEVER CONNECTED"

    current_time = now or datetime.utcnow()
    elapsed_seconds = (current_time - last_seen).total_seconds()

    if elapsed_seconds <= ONLINE_THRESHOLD_MINUTES * 60:
        return "ONLINE"

    return "OFFLINE"


def _serialize_computer(computer):
    return {
        "id": computer.id,
        "computer_name": computer.computer_name,
        "manufacturer": computer.manufacturer,
        "model": computer.model,
        "serial_number": computer.serial_number,
        "operating_system": computer.operating_system,
        "os_version": computer.os_version,
        "cpu": computer.cpu,
        "ram": computer.ram,
        "storage": computer.storage,
        "ip_address": computer.ip_address,
        "mac_address": computer.mac_address,
        "department_id": computer.department_id,
        "department_name": computer.department.name if computer.department else None,
        "assigned_user": computer.assigned_user,
        "location": computer.location,
        "status": computer.status,
        "purchase_date": computer.purchase_date.isoformat() if computer.purchase_date else None,
        "warranty": computer.warranty,
        "notes": computer.notes,
        "installed_software": computer.installed_software,
        "created_at": computer.created_at.isoformat() if computer.created_at else None,
        "updated_at": computer.updated_at.isoformat() if computer.updated_at else None,
        "last_seen": computer.last_seen.isoformat() if computer.last_seen else None,
        "connection": get_connection_status(computer.last_seen),
    }


def _serialize_pending_device(device):
    return {
        "id": device.id,
        "computer_name": device.computer_name,
        "manufacturer": device.manufacturer,
        "model": device.model,
        "serial_number": device.serial_number,
        "operating_system": device.operating_system,
        "os_version": device.os_version,
        "cpu": device.cpu,
        "ram": device.ram,
        "storage": device.storage,
        "ip_address": device.ip_address,
        "mac_address": device.mac_address,
        "installed_software": device.installed_software,
        "last_seen": device.last_seen.isoformat() if device.last_seen else None,
    }


def get_pending_devices(session):
    return [
        _serialize_pending_device(device)
        for device in session.query(PendingDevice)
        .order_by(PendingDevice.computer_name)
        .all()
    ]


def approve_pending_device(session, device_id, approval_data=None):
    device = session.get(PendingDevice, device_id)
    if device is None:
        raise LookupError(f"Pending device {device_id} was not found")

    existing = session.query(Computer).filter_by(
        computer_name=device.computer_name
    ).first()
    if existing is not None:
        raise ValueError(f"Computer '{device.computer_name}' already exists")

    approval_data = approval_data or {}
    department_id = approval_data.get("department_id")
    department = session.get(Department, department_id) if department_id else None
    if department_id and department is None:
        raise ValueError("The selected department does not exist")

    status = normalize_status(approval_data.get("status"))

    computer = Computer(
        computer_name=device.computer_name,
        manufacturer=device.manufacturer,
        model=device.model,
        serial_number=device.serial_number,
        operating_system=device.operating_system,
        os_version=device.os_version,
        cpu=device.cpu,
        ram=device.ram,
        storage=device.storage,
        ip_address=device.ip_address,
        mac_address=device.mac_address,
        installed_software=device.installed_software,
        last_seen=device.last_seen,
        department_id=department.id if department else None,
        assigned_user=approval_data.get("assigned_user") or None,
        location=approval_data.get("location") or None,
        status=status,
    )
    session.add(computer)
    session.delete(device)
    session.flush()
    return _serialize_computer(computer)


def _value(filters, key):
    value = filters.get(key)
    if isinstance(value, (list, tuple)):
        return value[0] if value else None
    return value


def get_computers(session, filters=None):
    query = session.query(Computer)
    filters = filters or {}
    status = _value(filters, "status")
    department_id = _value(filters, "department_id")
    search = _value(filters, "search")
    if status:
        query = query.filter(Computer.status == status)
    if department_id:
        query = query.filter(Computer.department_id == int(department_id))
    if search:
        term = search.lower()
        query = query.filter(
            (Computer.computer_name.ilike(f"%{term}%"))
            | (Computer.serial_number.ilike(f"%{term}%"))
            | (Computer.assigned_user.ilike(f"%{term}%"))
            | (Computer.ip_address.ilike(f"%{term}%"))
            | (Computer.mac_address.ilike(f"%{term}%"))
            | (Computer.manufacturer.ilike(f"%{term}%"))
        )
    return [_serialize_computer(item) for item in query.order_by(Computer.computer_name).all()]


def get_computer(session, computer_id):
    computer = session.get(Computer, computer_id)
    if computer is None:
        raise LookupError(f"Computer {computer_id} was not found")
    return computer


def get_computer_data(session, computer_id):
    computer = get_computer(session, computer_id)
    return _serialize_computer(computer)


def create_computer(session, payload):
    department_id = payload.get("department_id")
    dept = session.get(Department, department_id) if department_id else None
    if department_id and dept is None:
        raise ValueError("The selected department does not exist")
    computer = Computer(
        computer_name=(payload.get("computer_name") or "").strip(),
        manufacturer=payload.get("manufacturer"),
        model=payload.get("model"),
        serial_number=payload.get("serial_number"),
        operating_system=payload.get("operating_system"),
        os_version=payload.get("os_version"),
        cpu=payload.get("cpu"),
        ram=payload.get("ram"),
        storage=payload.get("storage"),
        ip_address=payload.get("ip_address"),
        mac_address=payload.get("mac_address"),
        department_id=dept.id if dept else None,
        assigned_user=payload.get("assigned_user"),
        location=payload.get("location") or "IT Storage",
        status=normalize_status(payload.get("status")),
        purchase_date=payload.get("purchase_date"),
        warranty=payload.get("warranty"),
        notes=payload.get("notes"),
    )
    session.add(computer)
    session.flush()
    log_action(session, computer, "Computer created", new_status=computer.status, description="Computer record created")
    return _serialize_computer(computer)


def update_computer(session, computer_id, payload):
    computer = get_computer(session, computer_id)
    previous_status = computer.status
    for field in [
        "computer_name", "manufacturer", "model", "serial_number", "operating_system",
        "os_version", "cpu", "ram", "storage", "ip_address", "mac_address",
        "assigned_user", "location", "status", "purchase_date", "warranty", "notes"
    ]:
        if field in payload:
            setattr(computer, field, payload[field])
    if "status" in payload:
        computer.status = normalize_status(payload["status"])
    if "department_id" in payload:
        computer.department_id = payload.get("department_id") or None
    session.flush()
    log_action(
        session,
        computer=computer,
        action="Computer edited",
        previous_status=previous_status,
        new_status=computer.status,
        description="Computer details updated",
    )
    return _serialize_computer(computer)


def delete_computer(session, computer_id):
    computer = get_computer(session, computer_id)
    log_action(session, computer, "Computer deleted", previous_status=computer.status, new_status="DELETED", description="Computer record removed")
    session.delete(computer)
    session.flush()
    return True


def move_to_stock(session, computer_id):
    computer = get_computer(session, computer_id)
    previous_status = computer.status
    computer.status = ComputerStatus.IN_STOCK
    computer.department_id = None
    computer.assigned_user = None
    computer.location = "IT Storage"
    session.flush()
    log_action(session, computer, "Computer moved to stock", previous_status=previous_status, new_status=computer.status, description="Computer returned to stock")
    return _serialize_computer(computer)


def reserve_computer(session, computer_id):
    computer = get_computer(session, computer_id)
    previous_status = computer.status
    computer.status = ComputerStatus.RESERVED
    session.flush()
    log_action(session, computer, "Computer reserved", previous_status=previous_status, new_status=computer.status, description="Computer reserved for future assignment")
    return _serialize_computer(computer)


def cancel_reservation(session, computer_id):
    computer = get_computer(session, computer_id)
    previous_status = computer.status
    computer.status = ComputerStatus.IN_STOCK
    session.flush()
    log_action(session, computer, "Reservation cancelled", previous_status=previous_status, new_status=computer.status, description="Reservation removed")
    return _serialize_computer(computer)


def assign_computer(session, computer_id, department_id, assigned_user, location):
    computer = get_computer(session, computer_id)
    department = session.get(Department, department_id)
    if department is None:
        raise ValueError("The selected department does not exist")
    previous_status = computer.status
    computer.department_id = department.id
    computer.assigned_user = assigned_user
    computer.location = location
    computer.status = ComputerStatus.ASSIGNED.value
    session.flush()
    log_action(session, computer, "Computer assigned", previous_status=previous_status, new_status=computer.status, description="Computer assigned to department")
    return _serialize_computer(computer)


def log_action(session, computer, action, previous_status=None, new_status=None, description=None, admin="admin"):
    log = AuditLog(
        computer_id=computer.id,
        action=action,
        previous_status=previous_status,
        new_status=new_status,
        department_id=computer.department_id,
        user=computer.assigned_user,
        admin=admin,
        description=description,
        created_at=datetime.utcnow(),
    )
    session.add(log)
