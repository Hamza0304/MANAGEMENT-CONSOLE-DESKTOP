
from database import engine, DATABASE_PATH
from sqlalchemy import text


def reset_device_test_data():
    print("=" * 60)
    print("MANAGEMENT CONSOLE - RESET DEVICE TEST DATA")
    print("=" * 60)

    print(f"\nDatabase:")
    print(DATABASE_PATH)

    with engine.begin() as connection:
        # Delete audit history related to computers and phones
        result = connection.execute(
            text("""
                DELETE FROM audit_logs
                WHERE computer_id IS NOT NULL
                   OR phone_id IS NOT NULL
            """)
        )

        print(f"\nDevice audit records deleted: {result.rowcount}")

        # Delete pending discovery devices
        connection.execute(
            text("DELETE FROM pending_devices")
        )

        # Delete phones
        connection.execute(
            text("DELETE FROM phones")
        )

        # Delete computers
        connection.execute(
            text("DELETE FROM computers")
        )

    print("\nReset completed successfully.")

    print("\nCleared:")
    print("  computers")
    print("  phones")
    print("  pending_devices")
    print("  computer/phone audit history")

    print("\nPreserved:")
    print("  departments")
    print("  users")
    print("  settings")
    print("  other application data")
    print("  audit logs unrelated to computers/phones")

    print("=" * 60)


if __name__ == "__main__":
    reset_device_test_data()

