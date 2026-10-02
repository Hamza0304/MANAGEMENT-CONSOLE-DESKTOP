from database import engine, DATABASE_PATH
from sqlalchemy import text

print("=" * 60)
print("DATABASE CHECK")
print("=" * 60)

print("Database path:")
print(DATABASE_PATH)

with engine.connect() as connection:
    print("\nCURRENT COUNTS:")

    for table in [
        "computers",
        "phones",
        "pending_devices",
        "audit_logs",
    ]:
        result = connection.execute(
            text(f"SELECT COUNT(*) FROM {table}")
        )
        count = result.scalar()
        print(f"{table}: {count}")

    print("\nAUDIT LOG ROWS:")

    result = connection.execute(
        text("""
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
    )

    rows = result.fetchall()

    for row in rows:
        print(row)

print("\nDone.")