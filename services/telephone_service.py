from datetime import datetime

import pandas as pd

from models import Department, Telephone, TelephoneStatus
from models.telephone import normalize_telephone_status


def _serialize_telephone(telephone):
    return {
        "id": telephone.id,
        "asset_tag": telephone.asset_tag,
        "telephone_number": telephone.telephone_number,
        "brand": telephone.brand,
        "model": telephone.model,
        "serial_number": telephone.serial_number,
        "mac_address": telephone.mac_address,
        "ip_address": telephone.ip_address,
        "department_id": telephone.department_id,
        "department_name": telephone.department.name if telephone.department else None,
        "assigned_user": telephone.assigned_user,
        "location": telephone.location,
        "status": telephone.status,
        "purchase_date": telephone.purchase_date.isoformat() if telephone.purchase_date else None,
        "notes": telephone.notes,
        "created_at": telephone.created_at.isoformat() if telephone.created_at else None,
        "updated_at": telephone.updated_at.isoformat() if telephone.updated_at else None,
    }


def _validate_unique_fields(session, payload, telephone_id=None):
    for field in ("asset_tag", "telephone_number", "serial_number"):
        value = (payload.get(field) or "").strip()
        if not value:
            continue
        query = session.query(Telephone).filter(getattr(Telephone, field) == value)
        if telephone_id is not None:
            query = query.filter(Telephone.id != telephone_id)
        if query.first() is not None:
            raise ValueError(f"{field.replace('_', ' ').title()} already exists")


def get_telephones(session, filters=None):
    query = session.query(Telephone)
    filters = filters or {}
    if filters.get("status"):
        query = query.filter(Telephone.status == filters["status"])
    if filters.get("department_id"):
        query = query.filter(Telephone.department_id == int(filters["department_id"]))
    if filters.get("search"):
        term = f"%{str(filters['search']).lower()}%"
        query = query.filter(
            (Telephone.asset_tag.ilike(term))
            | (Telephone.telephone_number.ilike(term))
            | (Telephone.serial_number.ilike(term))
            | (Telephone.brand.ilike(term))
            | (Telephone.model.ilike(term))
            | (Telephone.assigned_user.ilike(term))
        )
    if filters.get("location"):
        query = query.filter(Telephone.location.ilike(f"%{filters['location']}%"))
    if filters.get("assigned_user"):
        query = query.filter(Telephone.assigned_user.ilike(f"%{filters['assigned_user']}%"))
    return [_serialize_telephone(item) for item in query.order_by(Telephone.asset_tag).all()]


def get_telephone(session, telephone_id):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    return _serialize_telephone(telephone)


def create_telephone(session, payload):
    department_id = payload.get("department_id")
    department = session.get(Department, department_id) if department_id else None
    if department_id and department is None:
        raise ValueError("The selected department does not exist")
    _validate_unique_fields(session, payload)
    telephone = Telephone(
        asset_tag=(payload.get("asset_tag") or "").strip(),
        telephone_number=(payload.get("telephone_number") or "").strip(),
        brand=(payload.get("brand") or "").strip() or None,
        model=(payload.get("model") or "").strip() or None,
        serial_number=(payload.get("serial_number") or "").strip() or None,
        mac_address=(payload.get("mac_address") or "").strip() or None,
        ip_address=(payload.get("ip_address") or "").strip() or None,
        department_id=department.id if department else None,
        assigned_user=(payload.get("assigned_user") or "").strip() or None,
        location=(payload.get("location") or "").strip() or None,
        status=normalize_telephone_status(payload.get("status")),
        purchase_date=payload.get("purchase_date"),
        notes=(payload.get("notes") or "").strip() or None,
    )
    session.add(telephone)
    session.flush()
    return _serialize_telephone(telephone)


def update_telephone(session, telephone_id, payload):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    _validate_unique_fields(session, payload, telephone_id)
    for field in (
        "asset_tag", "telephone_number", "brand", "model", "serial_number",
        "mac_address", "ip_address", "assigned_user", "location", "notes",
    ):
        if field in payload:
            value = (payload.get(field) or "").strip()
            setattr(telephone, field, value or None)
    if "status" in payload:
        telephone.status = normalize_telephone_status(payload["status"])
    if "department_id" in payload:
        department_id = payload.get("department_id")
        if department_id and session.get(Department, department_id) is None:
            raise ValueError("The selected department does not exist")
        telephone.department_id = department_id or None
    if "purchase_date" in payload:
        telephone.purchase_date = payload["purchase_date"]
    session.flush()
    return _serialize_telephone(telephone)


def delete_telephone(session, telephone_id):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    session.delete(telephone)
    session.flush()
    return True


def move_telephone_to_stock(session, telephone_id):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    telephone.department_id = None
    telephone.assigned_user = None
    telephone.location = "IT Storage"
    telephone.status = TelephoneStatus.IN_STOCK.value
    session.flush()
    return _serialize_telephone(telephone)


def assign_telephone(session, telephone_id, department_id, assigned_user, location):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    department = session.get(Department, department_id)
    if department is None:
        raise ValueError("The selected department does not exist")
    telephone.department_id = department.id
    telephone.assigned_user = assigned_user
    telephone.location = location
    telephone.status = TelephoneStatus.ASSIGNED.value
    session.flush()
    return _serialize_telephone(telephone)


def reserve_telephone(session, telephone_id):
    telephone = session.get(Telephone, telephone_id)
    if telephone is None:
        raise LookupError(f"Telephone {telephone_id} was not found")
    telephone.status = TelephoneStatus.RESERVED.value
    session.flush()
    return _serialize_telephone(telephone)


def import_telephones(session, dataframe):
    columns = {
        str(column).strip().lower().replace(" ", "_"): column
        for column in dataframe.columns
    }
    if "asset_tag" not in columns or "telephone_number" not in columns:
        raise ValueError("The file must contain asset_tag and telephone_number columns")
    fields = (
        "asset_tag", "telephone_number", "brand", "model", "serial_number",
        "mac_address", "ip_address", "assigned_user", "location", "status", "notes",
    )
    created = 0
    skipped = 0
    for _, row in dataframe.iterrows():
        payload = {}
        for field in fields:
            source = columns.get(field)
            value = row[source] if source is not None else None
            payload[field] = None if value is None or pd.isna(value) else str(value).strip() or None
        if not payload["asset_tag"] or not payload["telephone_number"]:
            skipped += 1
            continue
        create_telephone(session, payload)
        created += 1
    session.flush()
    return {"created": created, "skipped": skipped, "rows": len(dataframe)}
