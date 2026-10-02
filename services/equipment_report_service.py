from models import Computer, Department, EquipmentAsset, EquipmentType, Phone, Telephone
from services.equipment_asset_service import _serialize as serialize_generic_asset


STATUS_VALUES = ("IN_STOCK", "ASSIGNED", "RESERVED", "MAINTENANCE", "LOST", "RETIRED")


def get_enabled_report_types(session):
    return [
        item
        for item in session.query(EquipmentType).order_by(EquipmentType.id).all()
        if item.enabled
    ]


def _serialize_common(item, equipment_type):
    return {
        "Equipment Type": equipment_type.name,
        "Asset Tag": getattr(item, "asset_tag", None) or getattr(item, "computer_name", None),
        "Name": (
            getattr(item, "computer_name", None)
            or getattr(item, "phone_number", None)
            or getattr(item, "telephone_number", None)
            or getattr(item, "name", None)
        ),
        "Brand": item.manufacturer if isinstance(item, Computer) else item.brand,
        "Model": item.model,
        "Serial Number": item.serial_number,
        "IP Address": item.ip_address,
        "MAC Address": item.mac_address,
        "Department": item.department.name if item.department else "",
        "Assigned User": item.assigned_user or "",
        "Location": item.location or "",
        "Status": item.status,
        "Notes": item.notes or "",
        "Created At": item.created_at.isoformat() if item.created_at else "",
        "Updated At": item.updated_at.isoformat() if item.updated_at else "",
    }


def get_equipment_report_rows(session, equipment_key, filters=None):
    filters = filters or {}
    equipment_type = session.query(EquipmentType).filter_by(key=equipment_key).first()
    if equipment_type is None:
        raise LookupError(f"Equipment type '{equipment_key}' was not found")

    if equipment_key == "computers":
        items = session.query(Computer).order_by(Computer.computer_name).all()
    elif equipment_key == "phones":
        items = session.query(Phone).order_by(Phone.asset_tag).all()
    elif equipment_key == "telephones":
        items = session.query(Telephone).order_by(Telephone.asset_tag).all()
    else:
        items = session.query(EquipmentAsset).filter_by(
            equipment_type_id=equipment_type.id
        ).order_by(EquipmentAsset.asset_tag).all()

    rows = []
    search = str(filters.get("search") or "").strip().lower()
    for item in items:
        row = (
            serialize_generic_asset(item)
            if isinstance(item, EquipmentAsset)
            else _serialize_common(item, equipment_type)
        )
        if isinstance(item, EquipmentAsset):
            row = {
                "Equipment Type": equipment_type.name,
                "Asset Tag": row["asset_tag"],
                "Name": row["name"] or "",
                "Brand": row["brand"] or "",
                "Model": row["model"] or "",
                "Serial Number": row["serial_number"] or "",
                "IP Address": row["ip_address"] or "",
                "MAC Address": row["mac_address"] or "",
                "Department": row["department_name"] or "",
                "Assigned User": row["assigned_user"] or "",
                "Location": row["location"] or "",
                "Status": row["status"] or "",
                "Notes": row["notes"] or "",
                "Created At": row["created_at"] or "",
                "Updated At": row["updated_at"] or "",
            }
        haystack = " ".join(str(value).lower() for value in row.values())
        if search and search not in haystack:
            continue
        if filters.get("status") and row["Status"] != filters["status"]:
            continue
        if filters.get("department_id"):
            department_id = getattr(item, "department_id", None)
            if department_id != int(filters["department_id"]):
                continue
        rows.append(row)
    return rows


def get_enabled_equipment_summary(session):
    summary = []
    for equipment_type in get_enabled_report_types(session):
        rows = get_equipment_report_rows(session, equipment_type.key)
        by_status = {status: 0 for status in STATUS_VALUES}
        by_department = {}
        for row in rows:
            by_status[row["Status"]] = by_status.get(row["Status"], 0) + 1
            department = row["Department"] or "Unassigned"
            by_department[department] = by_department.get(department, 0) + 1
        summary.append({
            "key": equipment_type.key,
            "name": equipment_type.name,
            "total": len(rows),
            "by_status": by_status,
            "by_department": by_department,
        })
    return summary


def get_enabled_department_summary(session):
    result = []
    for department in session.query(Department).order_by(Department.name).all():
        counts = {}
        for equipment_type in get_enabled_report_types(session):
            rows = get_equipment_report_rows(
                session,
                equipment_type.key,
                {"department_id": department.id},
            )
            if rows:
                counts[equipment_type.name] = len(rows)
        result.append({"Department": department.name, **counts})
    return result
