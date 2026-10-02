import pandas as pd

from models import Department, EquipmentAsset, EquipmentType
from services.equipment_type_service import EQUIPMENT_STATUSES


STATUSES = EQUIPMENT_STATUSES


def _type(session, equipment_key):
    equipment_type = session.query(EquipmentType).filter_by(key=equipment_key).first()
    if equipment_type is None:
        raise LookupError(f"Equipment type '{equipment_key}' was not found")
    return equipment_type


def _serialize(asset):
    return {
        "id": asset.id,
        "equipment_type_key": asset.equipment_type.key if asset.equipment_type else None,
        "equipment_type_name": asset.equipment_type.name if asset.equipment_type else None,
        "asset_tag": asset.asset_tag,
        "name": asset.name,
        "brand": asset.brand,
        "model": asset.model,
        "serial_number": asset.serial_number,
        "ip_address": asset.ip_address,
        "mac_address": asset.mac_address,
        "department_id": asset.department_id,
        "department_name": asset.department.name if asset.department else None,
        "assigned_user": asset.assigned_user,
        "location": asset.location,
        "status": asset.status,
        "purchase_date": asset.purchase_date.isoformat() if asset.purchase_date else None,
        "notes": asset.notes,
        "created_at": asset.created_at.isoformat() if asset.created_at else None,
        "updated_at": asset.updated_at.isoformat() if asset.updated_at else None,
    }


def _validate_unique(session, payload, equipment_type_id, asset_id=None):
    asset_tag = (payload.get("asset_tag") or "").strip()
    if asset_id is None and not asset_tag:
        raise ValueError("Asset tag is required")
    if not asset_tag:
        return
    query = session.query(EquipmentAsset).filter(
        EquipmentAsset.equipment_type_id == equipment_type_id,
        EquipmentAsset.asset_tag == asset_tag,
    )
    if asset_id is not None:
        query = query.filter(EquipmentAsset.id != asset_id)
    if query.first() is not None:
        raise ValueError("Asset tag already exists for this equipment type")


def get_equipment_assets(session, equipment_key, filters=None):
    equipment_type = _type(session, equipment_key)
    filters = filters or {}
    query = session.query(EquipmentAsset).filter(
        EquipmentAsset.equipment_type_id == equipment_type.id
    )
    if filters.get("status"):
        query = query.filter(EquipmentAsset.status == filters["status"])
    if filters.get("department_id"):
        query = query.filter(EquipmentAsset.department_id == int(filters["department_id"]))
    if filters.get("search"):
        term = f"%{str(filters['search']).lower()}%"
        query = query.filter(
            (EquipmentAsset.asset_tag.ilike(term))
            | (EquipmentAsset.name.ilike(term))
            | (EquipmentAsset.brand.ilike(term))
            | (EquipmentAsset.model.ilike(term))
            | (EquipmentAsset.serial_number.ilike(term))
            | (EquipmentAsset.assigned_user.ilike(term))
        )
    if filters.get("location"):
        query = query.filter(EquipmentAsset.location.ilike(f"%{filters['location']}%"))
    if filters.get("assigned_user"):
        query = query.filter(EquipmentAsset.assigned_user.ilike(f"%{filters['assigned_user']}%"))
    return [_serialize(asset) for asset in query.order_by(EquipmentAsset.asset_tag).all()]


def get_equipment_asset(session, equipment_key, asset_id):
    equipment_type = _type(session, equipment_key)
    asset = session.query(EquipmentAsset).filter_by(
        id=asset_id,
        equipment_type_id=equipment_type.id,
    ).first()
    if asset is None:
        raise LookupError(f"{equipment_key} asset {asset_id} was not found")
    return _serialize(asset)


def create_equipment_asset(session, equipment_key, payload):
    equipment_type = _type(session, equipment_key)
    _validate_unique(session, payload, equipment_type.id)
    department_id = payload.get("department_id")
    if department_id and session.get(Department, department_id) is None:
        raise ValueError("The selected department does not exist")
    status = str(payload.get("status") or "IN_STOCK").strip().upper().replace(" ", "_")
    if status not in STATUSES:
        raise ValueError(f"Invalid status: {payload.get('status')}")
    asset = EquipmentAsset(
        equipment_type_id=equipment_type.id,
        asset_tag=(payload.get("asset_tag") or "").strip(),
        name=(payload.get("name") or "").strip() or None,
        brand=(payload.get("brand") or "").strip() or None,
        model=(payload.get("model") or "").strip() or None,
        serial_number=(payload.get("serial_number") or "").strip() or None,
        ip_address=(payload.get("ip_address") or "").strip() or None,
        mac_address=(payload.get("mac_address") or "").strip() or None,
        department_id=department_id or None,
        assigned_user=(payload.get("assigned_user") or "").strip() or None,
        location=(payload.get("location") or "").strip() or None,
        status=status,
        purchase_date=payload.get("purchase_date"),
        notes=(payload.get("notes") or "").strip() or None,
    )
    session.add(asset)
    session.flush()
    return _serialize(asset)


def update_equipment_asset(session, equipment_key, asset_id, payload):
    equipment_type = _type(session, equipment_key)
    asset = session.query(EquipmentAsset).filter_by(
        id=asset_id,
        equipment_type_id=equipment_type.id,
    ).first()
    if asset is None:
        raise LookupError(f"{equipment_key} asset {asset_id} was not found")
    _validate_unique(session, payload, equipment_type.id, asset_id)
    for field in (
        "asset_tag", "name", "brand", "model", "serial_number", "ip_address",
        "mac_address", "assigned_user", "location", "notes",
    ):
        if field in payload:
            value = (payload.get(field) or "").strip()
            setattr(asset, field, value or None)
    if "department_id" in payload:
        department_id = payload.get("department_id")
        if department_id and session.get(Department, department_id) is None:
            raise ValueError("The selected department does not exist")
        asset.department_id = department_id or None
    if "status" in payload:
        status = str(payload["status"] or "IN_STOCK").strip().upper().replace(" ", "_")
        if status not in STATUSES:
            raise ValueError(f"Invalid status: {payload['status']}")
        asset.status = status
    if "purchase_date" in payload:
        asset.purchase_date = payload["purchase_date"]
    session.flush()
    return _serialize(asset)


def delete_equipment_asset(session, equipment_key, asset_id):
    equipment_type = _type(session, equipment_key)
    asset = session.query(EquipmentAsset).filter_by(
        id=asset_id,
        equipment_type_id=equipment_type.id,
    ).first()
    if asset is None:
        raise LookupError(f"{equipment_key} asset {asset_id} was not found")
    session.delete(asset)
    session.flush()


def assign_equipment_asset(session, equipment_key, asset_id, department_id, assigned_user, location):
    return update_equipment_asset(session, equipment_key, asset_id, {
        "department_id": department_id,
        "assigned_user": assigned_user,
        "location": location,
        "status": "ASSIGNED",
    })


def move_equipment_asset_to_stock(session, equipment_key, asset_id):
    return update_equipment_asset(session, equipment_key, asset_id, {
        "department_id": None,
        "assigned_user": None,
        "location": "IT Storage",
        "status": "IN_STOCK",
    })


def reserve_equipment_asset(session, equipment_key, asset_id):
    return update_equipment_asset(session, equipment_key, asset_id, {"status": "RESERVED"})


def import_equipment_assets(session, equipment_key, dataframe):
    columns = {
        str(column).strip().lower().replace(" ", "_"): column
        for column in dataframe.columns
    }
    if "asset_tag" not in columns:
        raise ValueError("The file must contain an asset_tag column")
    fields = (
        "asset_tag", "name", "brand", "model", "serial_number", "ip_address",
        "mac_address", "assigned_user", "location", "status", "notes",
    )
    created = 0
    skipped = 0
    for _, row in dataframe.iterrows():
        payload = {}
        for field in fields:
            source = columns.get(field)
            value = row[source] if source is not None else None
            payload[field] = None if value is None or pd.isna(value) else str(value).strip() or None
        if not payload["asset_tag"]:
            skipped += 1
            continue
        create_equipment_asset(session, equipment_key, payload)
        created += 1
    session.flush()
    return {"created": created, "skipped": skipped, "rows": len(dataframe)}
