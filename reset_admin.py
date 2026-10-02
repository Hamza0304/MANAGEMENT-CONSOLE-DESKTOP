from database import SessionLocal
from security import set_admin_password

session = SessionLocal()

try:
    set_admin_password(session, "admin", "Admin@1234")
    print("Admin password reset successfully.")
    print("Username: admin")
    print("Password: Admin@1234")
finally:
    session.close()
