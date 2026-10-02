from werkzeug.security import generate_password_hash, check_password_hash

from models.admin import Admin


def authenticate(session, username, password):
    admin = session.query(Admin).filter(
        Admin.username == username
    ).first()

    if admin is None:
        return None

    if not check_password_hash(admin.password_hash, password):
        return None

    return admin


def set_admin_password(session, username, new_password):
    admin = session.query(Admin).filter(
        Admin.username == username
    ).first()

    if admin is None:
        raise ValueError(f"Administrator '{username}' not found.")

    admin.password_hash = generate_password_hash(new_password)
    session.commit()

    return admin

def ensure_default_admin(session):
    admin = session.query(Admin).filter(
        Admin.username == "admin"
    ).first()

    if admin is None:
        admin = Admin(
            username="admin",
            password_hash=generate_password_hash("Admin@1234"),
            role="ADMIN",
        )
        session.add(admin)
        session.commit()

    return admin