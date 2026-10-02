from models import Department, EquipmentType
from models.equipment_type import DEFAULT_EQUIPMENT_TYPES

EQUIPMENT_STATUSES = (
    "IN_STOCK",
    "ASSIGNED",
    "RESERVED",
    "MAINTENANCE",
    "LOST",
    "RETIRED",
)


def get_equipment_filter_options(session):
    return {
        "departments": session.query(Department).order_by(Department.name).all(),
        "statuses": EQUIPMENT_STATUSES,
    }


SPECIAL_EQUIPMENT_PAGE_KEYS = {
    "computers": "computers",
    "phones": "phones",
    "telephones": "telephones",
}


def get_equipment_page_kind(equipment_key):
    return SPECIAL_EQUIPMENT_PAGE_KEYS.get(equipment_key, "equipment_type")


def create_equipment_type(session, key, name, description=None, enabled=False):
    normalized_key = str(key or "").strip().lower().replace(" ", "_")
    normalized_name = str(name or "").strip()
    if not normalized_key or not normalized_name:
        raise ValueError("Equipment type name and key are required")
    if not all(character.isalnum() or character == "_" for character in normalized_key):
        raise ValueError("Equipment type key may contain only letters, numbers, and underscores")
    if session.query(EquipmentType).filter_by(key=normalized_key).first() is not None:
        raise ValueError("Equipment type key already exists")
    equipment_type = EquipmentType(
        key=normalized_key,
        name=normalized_name,
        description=(description or "").strip() or None,
        enabled=bool(enabled),
    )
    session.add(equipment_type)
    session.flush()
    return equipment_type


def ensure_default_equipment_types(session):
    existing = {
        item.key: item
        for item in session.query(EquipmentType).all()
    }
    for key, name, description, enabled in DEFAULT_EQUIPMENT_TYPES:
        if key not in existing:
            session.add(
                EquipmentType(
                    key=key,
                    name=name,
                    description=description,
                    enabled=enabled,
                )
            )
    session.flush()


def get_equipment_types(session):
    ensure_default_equipment_types(session)
    return session.query(EquipmentType).order_by(EquipmentType.id).all()


def get_enabled_equipment_types(session):
    return [item for item in get_equipment_types(session) if item.enabled]


def set_equipment_type_enabled(session, key, enabled):
    equipment_type = session.query(EquipmentType).filter_by(key=key).first()
    if equipment_type is None:
        raise LookupError(f"Equipment type '{key}' was not found")
    equipment_type.enabled = bool(enabled)
    session.flush()
    return equipment_type
