
from database import engine, DATABASE_PATH
from sqlalchemy import text


def reset_test_devices():
    print("=" * 60)
    print("MANAGEMENT CONSOLE - DEVICE DATA RESET")
    print("=" * 60)

    print(f"\nDatabase being used:")
    print(DATABASE_PATH)

    print("\nClearing test device data...")

    with engine.begin() as connection:
        # Remove pending discovery devices
        connection.execute(text("DELETE FROM pending_devices"))

        # Remove phones
        connection.execute(text("DELETE FROM phones"))

        # Remove computers
        connection.execute(text("DELETE FROM computers"))

    print("\nDone.")
    print("  computers       -> cleared")
    print("  phones          -> cleared")
    print("  pending_devices -> cleared")

    print("\nOther data was NOT deleted:")
    print("  departments")
    print("  users")
    print("  settings")
    print("  other application data")

    print("=" * 60)


if __name__ == "__main__":
    reset_test_devices()

