import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from flask import Flask, request

logger = logging.getLogger(__name__)

from database import (
    APP_ROOT,
    DATABASE_PATH,
    DATABASE_URL,
    engine,
    initialize_database,
    session_scope,
)
from models import Computer, PendingDevice

app = Flask(__name__)

API_KEY = os.getenv("MANAGEMENT_CONSOLE_API_KEY", "").strip()


def authenticate_agent():
    if not API_KEY:
        print("WARNING: MANAGEMENT_CONSOLE_API_KEY is not configured")
        return False

    auth_header = request.headers.get("Authorization", "").strip()

    if not auth_header.startswith("Bearer "):
        return False

    provided_key = auth_header[7:].strip()
    return provided_key == API_KEY


def print_database_diagnostics():
    database_file = Path(engine.url.database).resolve() if engine.url.database else None
    print("API DATABASE DIAGNOSTICS")
    print("  DATABASE_URL:", DATABASE_URL)
    print("  DATABASE_PATH:", DATABASE_PATH)
    print("  APP_ROOT:", APP_ROOT)
    print("  absolute SQLite database path:", database_file)
    print("  current working directory:", Path.cwd())

    with engine.connect() as connection:
        computers = connection.exec_driver_sql(
            "SELECT id, computer_name FROM computers ORDER BY id"
        ).fetchall()
    print("  computers visible to API:", computers)


def _first_non_empty(values):
    if isinstance(values, list):
        for value in values:
            if value is not None and str(value).strip():
                return str(value).strip()
        return None
    return values if values is not None and str(values).strip() else None


def normalize_software_payload(value):
    if value is None:
        return []

    if isinstance(value, list):
        return value

    if isinstance(value, tuple):
        return list(value)

    if isinstance(value, dict):
        return [value]

    if isinstance(value, str):
        text = value.strip()
        if not text or text in {"[]", "null", "None"}:
            return []
        try:
            parsed = json.loads(text)
        except (TypeError, ValueError):
            return [{"name": text}] if text else []

        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]
        return []

    return []


@app.route("/api/agent", methods=["POST"])
def receive_agent():
    if not authenticate_agent():
        return {"success": False, "message": "Unauthorized"}, 401

    data = request.get_json(silent=True) or {}
    computer_name = str(data.get("computer_name") or "").strip()

    print("Computer name received:", computer_name)
    print_database_diagnostics()

    if not computer_name:
        return {"success": False, "message": "computer_name is required"}, 400

    cpu_info = data.get("cpu") or {}
    ram_info = data.get("ram") or {}
    disk_info = data.get("disk") or {}
    data_ip = _first_non_empty(data.get("ip_address") or data.get("ip_addresses"))
    data_mac = _first_non_empty(data.get("mac_address") or data.get("mac_addresses"))
    software_payload = normalize_software_payload(data.get("installed_software"))
    logger.info("[DEBUG] installed_software received for %s: %s entries", computer_name, len(software_payload))
    print(f"[DEBUG] installed_software received for {computer_name}: {len(software_payload)} entries")

    response = None

    with session_scope() as session:
        computer = session.query(Computer).filter_by(computer_name=computer_name).first()

        if computer is None:
            pending = session.query(PendingDevice).filter_by(computer_name=computer_name).first()
            if pending is None:
                pending = PendingDevice(computer_name=computer_name)
                session.add(pending)

            pending.manufacturer = data.get("manufacturer")
            pending.model = data.get("model")
            pending.serial_number = data.get("serial_number")
            pending.cpu = cpu_info.get("name")
            pending.ram = f"{ram_info['total_gb']} GB" if ram_info.get("total_gb") is not None else None
            pending.storage = f"{disk_info['total_gb']} GB" if disk_info.get("total_gb") is not None else None
            pending.operating_system = data.get("windows")
            pending.os_version = data.get("os_version")
            pending.ip_address = data_ip
            pending.mac_address = data_mac
            print(f"[DEBUG] saving installed_software to PendingDevice for {computer_name}: {len(software_payload)} entries")
            logger.info("[DEBUG] saving installed_software to PendingDevice for %s: %s entries", computer_name, len(software_payload))
            pending.installed_software = json.dumps(software_payload, ensure_ascii=False)
            pending.last_seen = datetime.utcnow()

            response = {
                "success": True,
                "pending": True,
                "message": "New device added to Pending Devices",
                "computer_name": computer_name,
            }, 202
        else:
            computer.manufacturer = data.get("manufacturer")
            computer.model = data.get("model")
            computer.cpu = cpu_info.get("name")

            if ram_info.get("total_gb") is not None:
                computer.ram = f"{ram_info['total_gb']} GB"
            if disk_info.get("total_gb") is not None:
                computer.storage = f"{disk_info['total_gb']} GB"

            computer.operating_system = data.get("windows")
            computer.os_version = data.get("os_version")
            computer.ip_address = data_ip
            computer.mac_address = data_mac
            computer.serial_number = data.get("serial_number")
            print(f"[DEBUG] saving installed_software to Computer for {computer_name}: {len(software_payload)} entries")
            logger.info("[DEBUG] saving installed_software to Computer for %s: %s entries", computer_name, len(software_payload))
            computer.installed_software = json.dumps(software_payload, ensure_ascii=False)
            computer.last_seen = datetime.utcnow()

            response = {
                "success": True,
                "message": "Computer information updated successfully",
                "computer_id": computer.id,
                "computer_name": computer.computer_name,
            }, 200

    return response


if __name__ == "__main__":
    initialize_database()
    print_database_diagnostics()
    from waitress import serve

    serve(app, host="0.0.0.0", port=5000)