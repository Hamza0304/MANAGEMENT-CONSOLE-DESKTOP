
import os
import sys
from pathlib import Path

# Add project root to Python path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


# ---------------------------------------------------------
# Server configuration
# ---------------------------------------------------------

def load_server_config():
    """
    Load server configuration.

    Development:
        config/server_config.json

    Frozen EXE:
        Config/server_config.json next to the EXE
    """

    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
        config_path = base_dir / "Config" / "server_config.json"
    else:
        config_path = PROJECT_ROOT / "config" / "server_config.json"

    if not config_path.exists():
        print(f"WARNING: Configuration file not found: {config_path}")
        return {}

    try:
        import json

        with config_path.open("r", encoding="utf-8") as file:
            config = json.load(file)

        return config

    except Exception as exc:
        print(f"ERROR: Could not load server configuration: {exc}")
        return {}


config = load_server_config()


# ---------------------------------------------------------
# Data directory
# ---------------------------------------------------------

data_dir = str(config.get("data_dir", "")).strip()

if data_dir:
    os.environ["MANAGEMENT_CONSOLE_DATA_DIR"] = data_dir


# ---------------------------------------------------------
# API key
# ---------------------------------------------------------

api_key = str(config.get("api_key", "")).strip()

if api_key:
    os.environ["MANAGEMENT_CONSOLE_API_KEY"] = api_key


# ---------------------------------------------------------
# Now import application modules
# ---------------------------------------------------------

from waitress import serve

from api.agent_api import app, print_database_diagnostics
from database import initialize_database


# ---------------------------------------------------------
# Start server
# ---------------------------------------------------------

if __name__ == "__main__":
    initialize_database()
    print_database_diagnostics()

    port = int(config.get("port", 5000))

    print("Management Console Server starting...")
    print(f"Server: http://0.0.0.0:{port}")

    serve(
        app,
        host="0.0.0.0",
        port=port,
    )
