from pathlib import Path
import sqlite3

db_path = Path("instance") / "management_console.db"

print("=" * 60)
print("INSTANCE DATABASE CHECK")
print("=" * 60)
print(f"Database: {db_path.resolve()}")

connection = sqlite3.connect(db_path)

for table in ["computers", "phones", "pending_devices", "audit_logs"]:
    try:
        cursor = connection.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        print(f"{table}: {count}")
    except Exception as exc:
        print(f"{table}: ERROR - {exc}")

print("\nAUDIT LOGS:")

try:
    cursor = connection.execute("""
        SELECT
            id,
            action,
            computer_id,
            phone_id,
            description,
            created_at
        FROM audit_logs
        ORDER BY id
    """)

    rows = cursor.fetchall()

    for row in rows:
        print(row)

except Exception as exc:
    print(f"Audit log error: {exc}")

connection.close()

print("\nDone.")