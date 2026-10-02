import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

from models import ComputerStatus

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PROFILE_FILE = PROJECT_ROOT / "instance" / "excel_mapping_profiles.json"


def _deduplicate_target_fields(fields):
    deduplicated = []
    seen_keys = {}

    for field in fields:
        key = (field.get("key") or "").strip()
        if not key:
            continue

        if key in seen_keys:
            existing = deduplicated[seen_keys[key]]
            existing["aliases"] = list(
                dict.fromkeys(
                    (existing.get("aliases") or []) + (field.get("aliases") or [])
                )
            )
            existing["required"] = existing.get("required", False) or field.get(
                "required", False
            )
            continue

        field_copy = dict(field)
        field_copy["aliases"] = list(dict.fromkeys(field.get("aliases") or []))
        deduplicated.append(field_copy)
        seen_keys[key] = len(deduplicated) - 1

    return deduplicated


TARGET_FIELDS = _deduplicate_target_fields([
    {
        "key": "computer_name",
        "label": "Computer Name",
        "required": True,
        "aliases": [
            "Computer Name",
            "PC Name",
            "PC",
            "Computer",
            "Hostname",
            "Host Name",
            "Device Name",
            "Machine Name",
            "Computer_Name",
            "PC_Name",
            "Host_Name",
            "Hostname",
            "Poste",
            "Nom PC",
            "Nom du PC",
            "Nom ordinateur",
        ],
    },
    {
        "key": "manufacturer",
        "label": "Manufacturer",
        "required": False,
        "aliases": [
            "Manufacturer",
            "Brand",
            "Make",
            "Vendor",
            "Marque",
            "Constructeur",
        ],
    },
    {
        "key": "model",
        "label": "Model",
        "required": False,
        "aliases": [
            "Model",
            "Model Name",
            "Device Model",
            "Modèle",
        ],
    },
    {
        "key": "serial_number",
        "label": "Serial Number",
        "required": False,
        "aliases": [
            "Serial Number",
            "Serial",
            "Serial No",
            "Serial #",
            "SN",
            "S/N",
            "Service Tag",
            "Numéro de série",
        ],
    },
    {
        "key": "operating_system",
        "label": "Operating System",
        "required": False,
        "aliases": [
            "Operating System",
            "OS",
            "System",
            "Operating System Name",
            "Système",
            "Système d'exploitation",
        ],
    },
    {
        "key": "os_version",
        "label": "OS Version",
        "required": False,
        "aliases": [
            "OS Version",
            "Windows Version",
            "System Version",
            "Version OS",
            "Version Windows",
        ],
    },
    {
        "key": "cpu",
        "label": "CPU",
        "required": False,
        "aliases": [
            "CPU",
            "Processor",
            "Processor Name",
            "CPU Model",
            "Processeur",
        ],
    },
    {
        "key": "ram",
        "label": "RAM",
        "required": False,
        "aliases": [
            "RAM",
            "Memory",
            "Memory Size",
            "RAM Size",
            "Mémoire",
        ],
    },
    {
        "key": "storage",
        "label": "Storage",
        "required": False,
        "aliases": [
            "Storage",
            "Disk",
            "Disk Size",
            "Hard Disk",
            "HDD",
            "SSD",
            "Capacity",
            "Stockage",
        ],
    },
    {
        "key": "ip_address",
        "label": "IP Address",
        "required": False,
        "aliases": [
            "IP",
            "IP Address",
            "IPv4",
            "Adresse IP",
        ],
    },
    {
        "key": "mac_address",
        "label": "MAC Address",
        "required": False,
        "aliases": [
            "MAC",
            "MAC Address",
            "MAC Addr",
            "Physical Address",
            "Adresse MAC",
        ],
    },
    {
        "key": "department",
        "label": "Department",
        "required": False,
        "aliases": [
            "Department",
            "Dept",
            "Department Name",
            "Service",
            "Division",
            "Département",
        ],
    },
    {
        "key": "assigned_user",
        "label": "Assigned User",
        "required": False,
        "aliases": [
            "Assigned User",
            "User",
            "Employee",
            "Employee Name",
            "Assigned To",
            "Person",
            "Staff",
            "Utilisateur",
            "Employé",
        ],
    },
    {
        "key": "location",
        "label": "Location",
        "required": False,
        "aliases": [
            "Location",
            "Room",
            "Office",
            "Building",
            "Site",
            "Position",
            "Emplacement",
            "Localisation",
        ],
    },
    {
        "key": "status",
        "label": "Status",
        "required": False,
        "aliases": [
            "Status",
            "State",
            "Computer Status",
            "Device Status",
            "État",
            "Statut",
        ],
    },
    {
        "key": "purchase_date",
        "label": "Purchase Date",
        "required": False,
        "aliases": [
            "Purchase Date",
            "Purchased",
            "Purchase",
            "Acquisition Date",
            "Date Purchased",
            "Date d'achat",
        ],
    },
    {
        "key": "warranty",
        "label": "Warranty",
        "required": False,
        "aliases": [
            "Warranty",
            "Warranty Status",
            "Warranty Expiry",
            "Warranty Period",
            "Garantie",
        ],
    },
    {
        "key": "notes",
        "label": "Notes",
        "required": False,
        "aliases": [
            "Notes",
            "Comment",
            "Comments",
            "Remark",
            "Remarks",
            "Description",
            "Observation",
        ],
    },
])

TARGET_FIELD_KEYS = [field["key"] for field in TARGET_FIELDS]
TARGET_FIELD_LABELS = {field["key"]: field["label"] for field in TARGET_FIELDS}

PROFILE_VERSION = 1


def normalize_column_name(value):
    if value is None:
        return ""

    text = str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.strip().lower()
    text = text.replace(" ", "")
    text = text.replace("_", "")
    text = text.replace("-", "")
    text = text.replace(".", "")
    text = text.replace("/", "")
    text = text.replace("&", "")
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


TARGET_FIELD_ALIAS_MAP = {
    field["key"]: [normalize_column_name(alias) for alias in field["aliases"]]
    for field in TARGET_FIELDS
}


def read_excel_dataframe(file_path):
    try:
        dataframe = pd.read_excel(file_path)
    except Exception as exc:  # pragma: no cover - UI handles message
        raise ValueError(f"Could not read the Excel file: {exc}") from exc

    if dataframe is None or dataframe.empty:
        raise ValueError("The Excel file is empty.")

    dataframe = dataframe.where(pd.notna(dataframe), None)
    return dataframe


def detect_auto_mapping(columns):
    columns = list(columns)
    normalized_columns = [normalize_column_name(column) for column in columns]
    detected = {}

    for source_column, column_norm in zip(columns, normalized_columns):
        best_target = None
        best_score = -1

        for field in TARGET_FIELDS:
            aliases = TARGET_FIELD_ALIAS_MAP.get(field["key"], [])
            score = score_match(column_norm, aliases)

            if score > best_score:
                best_target = field["key"]
                best_score = score

        if best_score >= 60:
            detected[source_column] = best_target

    return detected


def score_match(column_norm, aliases):
    if not aliases:
        return 0

    if column_norm in aliases:
        return 100

    for alias in aliases:
        if alias == column_norm:
            return 100
        if column_norm in alias or alias in column_norm:
            if len(alias) <= 2:
                return 70
            return 85

    return 0


def load_mapping_profiles():
    if not PROFILE_FILE.exists():
        return []

    try:
        with PROFILE_FILE.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return []

    if isinstance(data, list):
        return data

    return []


def save_mapping_profile(profile_name, mappings):
    profile_name = (profile_name or "").strip()
    if not profile_name:
        raise ValueError("Profile name is required.")

    profiles = load_mapping_profiles()

    profile_entry = {
        "name": profile_name,
        "mapping_version": PROFILE_VERSION,
        "created_date": datetime.utcnow().isoformat(),
        "updated_date": datetime.utcnow().isoformat(),
        "mappings": [
            {"excel_column": source_column, "target_field": target_field}
            for target_field, source_column in mappings.items()
            if source_column
        ],
    }

    existing_index = None
    for index, profile in enumerate(profiles):
        if profile.get("name", "").lower() == profile_name.lower():
            existing_index = index
            break

    if existing_index is None:
        profiles.append(profile_entry)
    else:
        profiles[existing_index] = {
            **profiles[existing_index],
            "updated_date": datetime.utcnow().isoformat(),
            "mapping_version": PROFILE_VERSION,
            "mappings": profile_entry["mappings"],
        }

    PROFILE_FILE.parent.mkdir(parents=True, exist_ok=True)

    with PROFILE_FILE.open("w", encoding="utf-8") as handle:
        json.dump(profiles, handle, ensure_ascii=False, indent=2)

    return profile_entry


def detect_mapping_profile(columns):
    profiles = load_mapping_profiles()
    normalized_columns = {normalize_column_name(column): column for column in columns if column}

    best_profile = None
    best_match_count = -1

    for profile in profiles:
        mappings = profile.get("mappings", [])
        if not mappings:
            continue

        target_fields = {item.get("target_field") for item in mappings if item.get("target_field")}
        if "computer_name" not in target_fields:
            continue

        match_count = 0
        for item in mappings:
            source_column = item.get("excel_column")
            if normalize_column_name(source_column) in normalized_columns:
                match_count += 1

        if match_count >= max(1, int(len(mappings) * 0.5)) and match_count > best_match_count:
            best_profile = profile
            best_match_count = match_count

    return best_profile


def build_mapping_from_profile(profile, columns):
    mapping = {}
    if not profile:
        return mapping

    normalized_lookup = {
        normalize_column_name(item.get("excel_column", "")): item.get("target_field")
        for item in profile.get("mappings", [])
        if item.get("excel_column") and item.get("target_field")
    }

    for column in columns:
        target = normalized_lookup.get(normalize_column_name(column))
        if target:
            mapping[column] = target

    return mapping


def build_preview_rows(dataframe, mapping_by_target, department_lookup, existing_computers, existing_serial_numbers):
    if not mapping_by_target or "computer_name" not in mapping_by_target:
        raise ValueError("Computer Name is required before importing.")

    preview_rows = []
    seen_names = set()
    seen_serials = set()
    required_targets = [field["key"] for field in TARGET_FIELDS if field["key"] in mapping_by_target]

    for index, row in dataframe.iterrows():
        payload = {
            "row_number": index + 2,
            "status": "valid",
            "messages": [],
            "warnings": [],
            "payload": {},
            "data": {},
        }

        for field in TARGET_FIELDS:
            if field["key"] not in mapping_by_target:
                continue

            source_column = mapping_by_target.get(field["key"])

            if source_column:
                raw_value = row.get(source_column)
                value = clean_cell(raw_value)
            else:
                value = None

            if field["key"] == "purchase_date":
                value = parse_purchase_date(value)

            payload["data"][field["key"]] = value
            payload["payload"][field["key"]] = value

        computer_name = payload["data"].get("computer_name")
        if not computer_name:
            payload["messages"].append("Computer Name is required before importing.")
        else:
            computer_key = normalize_column_name(computer_name)
            if computer_key in seen_names:
                payload["messages"].append(f"Duplicate computer name '{computer_name}' in Excel.")
            elif computer_key in existing_computers:
                payload["messages"].append(f"Computer '{computer_name}' already exists in the database.")
            else:
                seen_names.add(computer_key)

        serial_number = payload["data"].get("serial_number")
        if serial_number:
            serial_key = normalize_column_name(serial_number)
            if serial_key in seen_serials:
                payload["messages"].append(f"Duplicate serial number '{serial_number}' in Excel.")
            elif serial_key in existing_serial_numbers:
                payload["messages"].append(f"Serial number '{serial_number}' already exists in the database.")
            else:
                seen_serials.add(serial_key)

        department_name = payload["data"].get("department")
        department_id = None

        if department_name:
            department_lookup_key = normalize_column_name(department_name)
            department = department_lookup.get(department_lookup_key)

            if department:
                department_id = department["id"]
                payload["payload"]["department_id"] = department_id
            else:
                payload["warnings"].append(
                    f"Department '{department_name}' was not found; it will be created during import."
                )
                payload["payload"]["department_name"] = department_name

        payload["payload"]["department_id"] = department_id

        normalized_status = None
        status_value = payload["data"].get("status")
        if status_value:
            try:
                normalized_status = normalize_status_value(status_value)
                payload["payload"]["status"] = normalized_status
            except ValueError as exc:
                payload["messages"].append(str(exc))
                payload["payload"]["status"] = None
        else:
            if department_id:
                normalized_status = ComputerStatus.ASSIGNED.value
                payload["warnings"].append("Status was not provided; defaulted to ASSIGNED.")
            else:
                normalized_status = ComputerStatus.IN_STOCK.value
                payload["warnings"].append("Status was not provided; defaulted to IN_STOCK.")
            payload["payload"]["status"] = normalized_status

        if (
            normalized_status == ComputerStatus.ASSIGNED.value
            and not department_id
            and not payload.get("payload", {}).get("department_name")
        ):
            payload["messages"].append("An ASSIGNED computer must have a department.")

        payload["status"] = "error" if payload["messages"] else ("warning" if payload["warnings"] else "valid")
        payload["messages"] = payload["messages"] + payload["warnings"]
        payload["display_status"] = {
            "valid": "✓ Valid",
            "warning": "⚠ Warning",
            "error": "✗ Error",
        }[payload["status"]]

        preview_rows.append(payload)

    summary = {
        "total_rows": len(dataframe),
        "valid_rows": sum(1 for row in preview_rows if row["status"] == "valid"),
        "warning_rows": sum(1 for row in preview_rows if row["status"] == "warning"),
        "error_rows": sum(1 for row in preview_rows if row["status"] == "error"),
        "mapped_targets": required_targets,
    }

    return preview_rows, summary


def clean_cell(value):
    if value is None or pd.isna(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d") if isinstance(value, datetime) else value.date().isoformat()

    cleaned = str(value).strip()
    return cleaned or None


def parse_purchase_date(value):
    if value is None:
        return None

    try:
        return pd.to_datetime(value).date().isoformat()
    except Exception:
        return value


def normalize_status_value(value):
    if value is None:
        return ComputerStatus.IN_STOCK.value

    normalized = normalize_column_name(value)
    aliases = {
        "instock": ComputerStatus.IN_STOCK.value,
        "stock": ComputerStatus.IN_STOCK.value,
        "available": ComputerStatus.IN_STOCK.value,
        "active": ComputerStatus.IN_STOCK.value,
        "assigned": ComputerStatus.ASSIGNED.value,
        "assignee": ComputerStatus.ASSIGNED.value,
        "assigne": ComputerStatus.ASSIGNED.value,
        "reserve": ComputerStatus.RESERVED.value,
        "reserved": ComputerStatus.RESERVED.value,
        "maintenance": ComputerStatus.MAINTENANCE.value,
        "retired": ComputerStatus.RETIRED.value,
        "retire": ComputerStatus.RETIRED.value,
        "retiree": ComputerStatus.RETIRED.value,
        "inactive": ComputerStatus.RETIRED.value,
    }

    mapped = aliases.get(normalized)
    if mapped:
        return mapped

    raise ValueError(f"Invalid status '{value}'.")
