from datetime import datetime

from models import AuditLog, Department, Phone, PhoneStatus
from models.phone import normalize_phone_status


def _serialize_phone(phone):
    return {
        "id": phone.id,
        "asset_tag": phone.asset_tag,
        "phone_number": phone.phone_number,
        "brand": phone.brand,
        "model": phone.model,
        "imei_1": phone.imei_1,
        "imei_2": phone.imei_2,
        "serial_number": phone.serial_number,
        "operating_system": phone.operating_system,
        "os_version": phone.os_version,
        "mac_address": phone.mac_address,
        "ip_address": phone.ip_address,
        "department_id": phone.department_id,
        "department_name": phone.department.name if phone.department else None,
        "assigned_user": phone.assigned_user,
        "location": phone.location,
        "status": phone.status,
        "purchase_date": phone.purchase_date.isoformat() if phone.purchase_date else None,
        "warranty_expiry": phone.warranty_expiry.isoformat() if phone.warranty_expiry else None,
        "notes": phone.notes,
        "created_at": phone.created_at.isoformat() if phone.created_at else None,
        "updated_at": phone.updated_at.isoformat() if phone.updated_at else None,
    }


def _value(filters, key):
    value = filters.get(key)
    if isinstance(value, (list, tuple)):
        return value[0] if value else None
    return value


def _validate_unique_fields(session, payload, phone_id=None):
    asset_tag = (payload.get("asset_tag") or "").strip()
    if asset_tag:
        existing = session.query(Phone).filter(Phone.asset_tag == asset_tag)
        if phone_id is not None:
            existing = existing.filter(Phone.id != phone_id)
        if existing.first():
            raise ValueError("Asset tag already exists")

    phone_number = (payload.get("phone_number") or "").strip()
    if phone_number:
        existing = session.query(Phone).filter(Phone.phone_number == phone_number)
        if phone_id is not None:
            existing = existing.filter(Phone.id != phone_id)
        if existing.first():
            raise ValueError("Phone number already exists")

    for field_name in ["imei_1", "imei_2", "serial_number"]:
        value = (payload.get(field_name) or "").strip()
        if not value:
            continue
        existing = session.query(Phone).filter(getattr(Phone, field_name) == value)
        if phone_id is not None:
            existing = existing.filter(Phone.id != phone_id)
        if existing.first():
            if field_name == "imei_1":
                raise ValueError("IMEI 1 already exists")
            if field_name == "imei_2":
                raise ValueError("IMEI 2 already exists")
            raise ValueError("Serial number already exists")


def get_phones(session, filters=None):
    query = session.query(Phone)
    filters = filters or {}
    status = _value(filters, "status")
    department_id = _value(filters, "department_id")
    search = _value(filters, "search")
    if status:
        query = query.filter(Phone.status == status)
    if department_id:
        query = query.filter(Phone.department_id == int(department_id))
    if search:
        term = search.lower()
        query = query.filter(
            (Phone.asset_tag.ilike(f"%{term}%"))
            | (Phone.phone_number.ilike(f"%{term}%"))
            | (Phone.imei_1.ilike(f"%{term}%"))
            | (Phone.imei_2.ilike(f"%{term}%"))
            | (Phone.serial_number.ilike(f"%{term}%"))
            | (Phone.brand.ilike(f"%{term}%"))
            | (Phone.model.ilike(f"%{term}%"))
            | (Phone.assigned_user.ilike(f"%{term}%"))
        )
    return [_serialize_phone(item) for item in query.order_by(Phone.asset_tag).all()]


def get_phone(session, phone_id):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")
    return _serialize_phone(phone)


def create_phone(session, payload):
    department_id = payload.get("department_id")
    dept = session.get(Department, department_id) if department_id else None
    if department_id and dept is None:
        raise ValueError("The selected department does not exist")

    _validate_unique_fields(session, payload)

    phone = Phone(
        asset_tag=(payload.get("asset_tag") or "").strip(),
        phone_number=(payload.get("phone_number") or "").strip(),
        brand=(payload.get("brand") or "").strip() or None,
        model=(payload.get("model") or "").strip() or None,
        imei_1=(payload.get("imei_1") or "").strip() or None,
        imei_2=(payload.get("imei_2") or "").strip() or None,
        serial_number=(payload.get("serial_number") or "").strip() or None,
        operating_system=(payload.get("operating_system") or "").strip() or None,
        os_version=(payload.get("os_version") or "").strip() or None,
        mac_address=(payload.get("mac_address") or "").strip() or None,
        ip_address=(payload.get("ip_address") or "").strip() or None,
        department_id=dept.id if dept else None,
        assigned_user=(payload.get("assigned_user") or "").strip() or None,
        location=(payload.get("location") or "").strip() or None,
        status=normalize_phone_status(payload.get("status")),
        purchase_date=payload.get("purchase_date"),
        warranty_expiry=payload.get("warranty_expiry"),
        notes=(payload.get("notes") or "").strip() or None,
    )

    session.add(phone)
    session.flush()
    _log_action(session, phone, "Phone created", new_status=phone.status, description="Phone record created")
    return _serialize_phone(phone)


def update_phone(session, phone_id, payload):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")

    previous_status = phone.status

    _validate_unique_fields(session, payload, phone_id=phone.id)

    for field in [
        "asset_tag", "phone_number", "brand", "model", "imei_1", "imei_2", "serial_number",
        "operating_system", "os_version", "mac_address", "ip_address", "assigned_user", "location",
        "status", "purchase_date", "warranty_expiry", "notes",
    ]:
        if field in payload:
            if field in {"asset_tag", "phone_number", "brand", "model", "imei_1", "imei_2", "serial_number",
                         "operating_system", "os_version", "mac_address", "ip_address", "assigned_user", "location",
                         "notes"}:
                value = (payload.get(field) or "").strip() if field != "notes" else (payload.get(field) or "").strip()
                setattr(phone, field, value or None)
            else:
                setattr(phone, field, payload[field])

    if "status" in payload:
        phone.status = normalize_phone_status(payload["status"])

    if "department_id" in payload:
        department_id = payload.get("department_id")
        if department_id:
            dept = session.get(Department, department_id)
            if dept is None:
                raise ValueError("The selected department does not exist")
            phone.department_id = dept.id
        else:
            phone.department_id = None

    session.flush()
    _log_action(
        session,
        phone,
        "Phone updated",
        previous_status=previous_status,
        new_status=phone.status,
        description="Phone details updated",
    )
    return _serialize_phone(phone)


def delete_phone(session, phone_id):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")

    _log_action(session, phone, "Phone deleted", previous_status=phone.status, new_status="DELETED", description="Phone record removed")
    session.delete(phone)
    session.flush()
    return True


def assign_phone(session, phone_id, department_id, assigned_user, location):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")

    department = session.get(Department, department_id)
    if department is None:
        raise ValueError("The selected department does not exist")

    previous_status = phone.status
    phone.department_id = department.id
    phone.assigned_user = assigned_user
    phone.location = location
    phone.status = PhoneStatus.ASSIGNED.value
    session.flush()
    _log_action(
        session,
        phone,
        "Phone assigned",
        previous_status=previous_status,
        new_status=phone.status,
        description="Phone assigned to department",
    )
    return _serialize_phone(phone)


def move_phone_to_stock(session, phone_id):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")

    previous_status = phone.status
    phone.department_id = None
    phone.assigned_user = None
    phone.location = "IT Storage"
    phone.status = PhoneStatus.IN_STOCK.value
    session.flush()
    _log_action(
        session,
        phone,
        "Phone moved to stock",
        previous_status=previous_status,
        new_status=phone.status,
        description="Phone returned to stock",
    )
    return _serialize_phone(phone)


def reserve_phone(session, phone_id):
    phone = session.get(Phone, phone_id)
    if phone is None:
        raise LookupError(f"Phone {phone_id} was not found")

    previous_status = phone.status
    phone.status = PhoneStatus.RESERVED.value
    session.flush()
    _log_action(
        session,
        phone,
        "Phone reserved",
        previous_status=previous_status,
        new_status=phone.status,
        description="Phone reserved for future assignment",
    )
    return _serialize_phone(phone)


def get_phone_stats(session):
    phones = session.query(Phone).all()
    total_phones = len(phones)
    assigned = sum(1 for phone in phones if phone.status == PhoneStatus.ASSIGNED.value)
    in_stock = sum(1 for phone in phones if phone.status == PhoneStatus.IN_STOCK.value)
    reserved = sum(1 for phone in phones if phone.status == PhoneStatus.RESERVED.value)
    maintenance = sum(1 for phone in phones if phone.status == PhoneStatus.MAINTENANCE.value)
    lost = sum(1 for phone in phones if phone.status == PhoneStatus.LOST.value)
    retired = sum(1 for phone in phones if phone.status == PhoneStatus.RETIRED.value)

    department_rows = []
    for department in session.query(Department).order_by(Department.name).all():
        department_rows.append(
            {
                "department": department.name,
                "phones": sum(1 for phone in department.phones),
                "computers": sum(1 for computer in department.computers),
                "total_assets": len(department.computers) + len(department.phones),
            }
        )

    return {
        "total_phones": total_phones,
        "assigned": assigned,
        "in_stock": in_stock,
        "reserved": reserved,
        "maintenance": maintenance,
        "lost": lost,
        "retired": retired,
        "department_totals": department_rows,
    }


def _log_action(session, phone, action, previous_status=None, new_status=None, description=None, admin="admin"):
    log = AuditLog(
        phone_id=phone.id,
        action=action,
        previous_status=previous_status,
        new_status=new_status,
        department_id=phone.department_id,
        user=phone.assigned_user,
        admin=admin,
        description=description,
        created_at=datetime.utcnow(),
    )
    session.add(log)
