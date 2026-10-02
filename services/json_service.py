import json

from models import Computer, Department, normalize_status


def import_json_file(session, file_storage):
    try:
        payload = json.load(file_storage)
    except Exception as exc:  # pragma: no cover
        return {"error": f"Invalid JSON file: {exc}"}

    if not isinstance(payload, list):
        return {"error": "JSON import expects a list of computer records."}

    created = 0
    for item in payload:
        name = (item.get("computer_name") or "").strip()
        if not name:
            continue
        dept_name = (item.get("department") or "").strip()
        dept = session.query(Department).filter_by(name=dept_name).first() if dept_name else None
        if dept_name and not dept:
            dept = Department(name=dept_name, description="Created during JSON import")
            session.add(dept)
            session.flush()

        computer = session.query(Computer).filter_by(computer_name=name).first()
        if computer is None:
            computer = Computer(
                computer_name=name,
                manufacturer=item.get("manufacturer"),
                model=item.get("model"),
                serial_number=item.get("serial_number"),
                operating_system=item.get("operating_system"),
                os_version=item.get("os_version"),
                cpu=item.get("cpu"),
                ram=item.get("ram"),
                storage=item.get("storage"),
                ip_address=item.get("ip_address"),
                mac_address=item.get("mac_address"),
                department_id=dept.id if dept else None,
                assigned_user=item.get("assigned_user"),
                location=item.get("location") or "IT Storage",
                status=normalize_status(item.get("status")),
                warranty=item.get("warranty"),
                notes=item.get("notes"),
            )
            session.add(computer)
            created += 1
    session.flush()
    return {"status": "ok", "created": created, "rows": len(payload)}
