from models import Department
from services.equipment_type_service import get_enabled_equipment_types


def _department_equipment_sections(session, department):
    sections = []
    enabled_types = get_enabled_equipment_types(session)
    source_records = {
        "computers": [
            {
                "id": item.id,
                "asset_tag": item.computer_name,
                "name": item.computer_name,
                "brand": item.manufacturer,
                "model": item.model,
                "serial_number": item.serial_number,
                "ip_address": item.ip_address,
                "status": item.status,
                "assigned_user": item.assigned_user,
                "location": item.location,
            }
            for item in department.computers
        ],
        "phones": [
            {
                "id": item.id,
                "asset_tag": item.asset_tag,
                "name": item.phone_number,
                "brand": item.brand,
                "model": item.model,
                "serial_number": item.serial_number,
                "ip_address": item.ip_address,
                "status": item.status,
                "assigned_user": item.assigned_user,
                "location": item.location,
            }
            for item in department.phones
        ],
        "telephones": [
            {
                "id": item.id,
                "asset_tag": item.asset_tag,
                "name": item.telephone_number,
                "brand": item.brand,
                "model": item.model,
                "serial_number": item.serial_number,
                "ip_address": item.ip_address,
                "status": item.status,
                "assigned_user": item.assigned_user,
                "location": item.location,
            }
            for item in department.telephones
        ],
    }
    for item in department.equipment_assets:
        source_records.setdefault(item.equipment_type.key, []).append(
            {
                "id": item.id,
                "asset_tag": item.asset_tag,
                "name": item.name,
                "brand": item.brand,
                "model": item.model,
                "serial_number": item.serial_number,
                "ip_address": item.ip_address,
                "status": item.status,
                "assigned_user": item.assigned_user,
                "location": item.location,
            }
        )
    for equipment_type in enabled_types:
        records = source_records.get(equipment_type.key, [])
        if records:
            sections.append(
                {
                    "key": equipment_type.key,
                    "name": equipment_type.name,
                    "records": records,
                }
            )
    return sections


def get_departments(session):
    departments = session.query(Department).order_by(Department.name).all()
    return [
        {
            "id": dept.id,
            "name": dept.name,
            "description": dept.description,
            "computer_count": len(dept.computers),
            "phone_count": len(dept.phones),
            "telephone_count": len(dept.telephones),
        }
        for dept in departments
    ]


def get_department(session, department_id):
    dept = session.get(Department, department_id)
    if dept is None:
        raise LookupError(f"Department {department_id} was not found")
    result = {
        "id": dept.id,
        "name": dept.name,
        "description": dept.description,
        "computer_count": len(dept.computers),
        "phone_count": len(dept.phones),
        "telephone_count": len(dept.telephones),
        "computers": [
            {
                "id": computer.id,
                "computer_name": computer.computer_name,
                "status": computer.status,
                "assigned_user": computer.assigned_user,
                "location": computer.location,
            }
            for computer in dept.computers
        ],
        "phones": [
            {
                "id": phone.id,
                "asset_tag": phone.asset_tag,
                "phone_number": phone.phone_number,
                "status": phone.status,
                "assigned_user": phone.assigned_user,
                "location": phone.location,
                "brand": phone.brand,
                "model": phone.model,
            }
            for phone in dept.phones
        ],
        "telephones": [
            {
                "id": telephone.id,
                "asset_tag": telephone.asset_tag,
                "telephone_number": telephone.telephone_number,
                "status": telephone.status,
                "assigned_user": telephone.assigned_user,
                "location": telephone.location,
                "brand": telephone.brand,
                "model": telephone.model,
            }
            for telephone in dept.telephones
        ],
        "equipment_assets": [
            {
                "id": asset.id,
                "equipment_type_key": asset.equipment_type.key,
                "equipment_type_name": asset.equipment_type.name,
                "asset_tag": asset.asset_tag,
                "name": asset.name,
                "brand": asset.brand,
                "model": asset.model,
                "serial_number": asset.serial_number,
                "ip_address": asset.ip_address,
                "status": asset.status,
                "assigned_user": asset.assigned_user,
                "location": asset.location,
            }
            for asset in dept.equipment_assets
        ],
    }
    result["equipment_sections"] = _department_equipment_sections(session, dept)
    return result


def create_department(session, payload):
    name = (payload.get("name") or "").strip()
    if not name:
        raise ValueError("Department name is required")
    dept = Department(name=name, description=payload.get("description"))
    session.add(dept)
    session.flush()
    return get_department(session, dept.id)


def update_department(session, department_id, payload):
    dept = session.get(Department, department_id)
    if dept is None:
        raise LookupError(f"Department {department_id} was not found")
    if "name" in payload:
        dept.name = (payload["name"] or "").strip()
    if "description" in payload:
        dept.description = payload.get("description")
    session.flush()
    return get_department(session, dept.id)


def delete_department(session, department_id):
    dept = session.get(Department, department_id)
    if dept is None:
        raise LookupError(f"Department {department_id} was not found")
    if dept.computers:
        raise ValueError("Move or delete the department's computers before deleting it")
    session.delete(dept)
    session.flush()
    return True
