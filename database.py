import os
import sys
from contextlib import contextmanager
from pathlib import Path


from sqlalchemy import create_engine, text
from sqlalchemy.orm import declarative_base, sessionmaker


PROJECT_ROOT = Path(__file__).resolve().parent

# ---------------------------------------------------------
# Application / Data directories
# ---------------------------------------------------------
#
# DEVELOPMENT:
#   Uses the project directory.
#
# PRODUCTION:
#   The installer/service can define:
#
#   MANAGEMENT_CONSOLE_DATA_DIR=D:\Management Console\Data
#
# If no custom data directory is configured, we keep the
# existing behavior to avoid moving or losing existing data.
# ---------------------------------------------------------

CUSTOM_DATA_DIR = os.environ.get("MANAGEMENT_CONSOLE_DATA_DIR", "").strip()

if CUSTOM_DATA_DIR:
    APP_ROOT = Path(CUSTOM_DATA_DIR).expanduser().resolve()
elif getattr(sys, "frozen", False):
    # Keep the current frozen-app behavior for compatibility.
    APP_ROOT = (
        Path(os.environ.get("LOCALAPPDATA", Path.home()))
        / "MANAGEMENT CONSOLE"
    )
else:
    APP_ROOT = PROJECT_ROOT

APP_ROOT.mkdir(parents=True, exist_ok=True)

DATABASE_PATH = APP_ROOT / "management_console.db"

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    f"sqlite:///{DATABASE_PATH}",
)

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, future=True, pool_pre_ping=True, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, expire_on_commit=False)
Base = declarative_base()


def test_connection():
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True


def _ensure_phones_table():
    if not DATABASE_URL.startswith("sqlite"):
        return

    with engine.begin() as connection:
        table_exists = connection.execute(
            text("SELECT 1 FROM sqlite_master WHERE type='table' AND name='phones'")
        ).fetchone()

        if table_exists is None:
            connection.execute(
                text(
                    """
                    CREATE TABLE phones (
                        id INTEGER NOT NULL PRIMARY KEY,
                        asset_tag VARCHAR(150) NOT NULL UNIQUE,
                        phone_number VARCHAR(50) NOT NULL UNIQUE,
                        brand VARCHAR(100),
                        model VARCHAR(150),
                        imei_1 VARCHAR(50) UNIQUE,
                        imei_2 VARCHAR(50) UNIQUE,
                        serial_number VARCHAR(150) UNIQUE,
                        operating_system VARCHAR(100),
                        os_version VARCHAR(100),
                        mac_address VARCHAR(50),
                        ip_address VARCHAR(50),
                        department_id INTEGER,
                        assigned_user VARCHAR(150),
                        location VARCHAR(150),
                        status VARCHAR(30) NOT NULL,
                        purchase_date DATE,
                        warranty_expiry DATE,
                        notes TEXT,
                        created_at DATETIME,
                        updated_at DATETIME,
                        FOREIGN KEY(department_id) REFERENCES departments(id)
                    )
                    """
                )
            )



def _ensure_audit_logs_phone_id():
    if not DATABASE_URL.startswith("sqlite"):
        return

    with engine.begin() as connection:
        table_exists = connection.execute(
            text(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='audit_logs'"
            )
        ).fetchone()

        if table_exists is None:
            return

        columns = connection.execute(
            text("PRAGMA table_info(audit_logs)")
        ).fetchall()

        column_names = {column[1] for column in columns}

        if "phone_id" not in column_names:
            connection.execute(
                text(
                    """
                    ALTER TABLE audit_logs
                    ADD COLUMN phone_id INTEGER
                    """
                )
            )


def _ensure_computers_last_seen():
    if not DATABASE_URL.startswith("sqlite"):
        return

    with engine.begin() as connection:
        table_exists = connection.execute(
            text(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='computers'"
            )
        ).fetchone()

        if table_exists is None:
            return

        columns = connection.execute(
            text("PRAGMA table_info(computers)")
        ).fetchall()

        if "last_seen" not in {column[1] for column in columns}:
            connection.execute(
                text(
                    "ALTER TABLE computers ADD COLUMN last_seen DATETIME"
                )
            )


def _ensure_computers_installed_software():
    if not DATABASE_URL.startswith("sqlite"):
        return

    with engine.begin() as connection:
        table_exists = connection.execute(
            text(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='computers'"
            )
        ).fetchone()

        if table_exists is None:
            return

        columns = connection.execute(
            text("PRAGMA table_info(computers)")
        ).fetchall()

        if "installed_software" not in {column[1] for column in columns}:
            connection.execute(
                text(
                    "ALTER TABLE computers ADD COLUMN installed_software TEXT"
                )
            )


def _ensure_pending_devices_os_version():
    if not DATABASE_URL.startswith("sqlite"):
        return

    with engine.begin() as connection:
        table_exists = connection.execute(
            text(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='pending_devices'"
            )
        ).fetchone()

        if table_exists is None:
            return

        columns = connection.execute(
            text("PRAGMA table_info(pending_devices)")
        ).fetchall()

        if "os_version" not in {column[1] for column in columns}:
            connection.execute(
                text(
                    "ALTER TABLE pending_devices ADD COLUMN os_version VARCHAR(100)"
                )
            )


def initialize_database():
    import models
    from services.equipment_type_service import ensure_default_equipment_types

    Base.metadata.create_all(bind=engine)
    _ensure_phones_table()
    _ensure_audit_logs_phone_id()
    _ensure_computers_last_seen()
    _ensure_computers_installed_software()
    _ensure_pending_devices_os_version()
    with SessionLocal() as session:
        ensure_default_equipment_types(session)
        session.commit()


@contextmanager
def session_scope():
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
