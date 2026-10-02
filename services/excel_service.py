import pandas as pd
from models import Computer, Department, normalize_status


def import_excel_file(session, file_storage):
    try:
        df = pd.read_excel(file_storage)
    except Exception as exc:  # pragma: no cover
        return {"error": f"Invalid Excel file: {exc}"}

    required = {"computer_name", "department"}
    missing = required - set(df.columns.astype(str))
    if missing:
        return {"error": f"Missing required columns: {sorted(missing)}"}

    created = 0
    warnings = []
    for _, row in df.iterrows():
        name = str(row.get("computer_name") or "").strip()
        if not name:
            warnings.append("Skipped empty computer name")
            continue
        dept_name = str(row.get("department") or "").strip()
        dept = session.query(Department).filter_by(name=dept_name).first()
        if not dept:
            dept = Department(name=dept_name, description="Created during Excel import")
            session.add(dept)
            session.flush()
        computer = session.query(Computer).filter_by(computer_name=name).first()
        if computer is None:
            computer = Computer(
                computer_name=name,
                manufacturer=str(row.get("manufacturer") or "") or None,
                model=str(row.get("model") or "") or None,
                serial_number=str(row.get("serial_number") or "") or None,
                operating_system=str(row.get("operating_system") or "") or None,
                os_version=str(row.get("os_version") or "") or None,
                cpu=str(row.get("cpu") or "") or None,
                ram=str(row.get("ram") or "") or None,
                storage=str(row.get("storage") or "") or None,
                ip_address=str(row.get("ip_address") or "") or None,
                mac_address=str(row.get("mac_address") or "") or None,
                department_id=dept.id,
                assigned_user=str(row.get("assigned_user") or "") or None,
                location=str(row.get("location") or "IT Storage") or "IT Storage",
                status=normalize_status(row.get("status")),
                warranty=str(row.get("warranty") or "") or None,
                notes=str(row.get("notes") or "") or None,
            )
            session.add(computer)
            created += 1
    session.flush()
    return {"status": "ok", "created": created, "warnings": warnings, "rows": len(df)}
